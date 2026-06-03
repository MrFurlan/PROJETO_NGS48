Imports System.Data
Imports System.Collections
Imports System.IO
Imports System.Text
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class CadastroDePlacas
    Inherits BasePage

#Region "Events"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Expedicao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("CadastroDePlacas", "ACESSAR") Then
                    CarregarTipoVeiculo()
                    CarregarViaTransporte()
                    Limpar()
                    CarregarEstados()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Expedicao.aspx")
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregarEstados()
        ddl.Carregar(ddlEstadoPlaca1, CarregarDDL.Tabela.EstadoERegiao, "", True)
        ddl.Carregar(ddlEstadoPlaca2, CarregarDDL.Tabela.EstadoERegiao, "", True)
        ddl.Carregar(ddlEstadoPlaca3, CarregarDDL.Tabela.EstadoERegiao, "", True)
        ddl.Carregar(ddlEstadoPlaca4, CarregarDDL.Tabela.EstadoERegiao, "", True)
    End Sub

    Protected Sub ddlEstadoPlaca1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlEstadoPlaca1.SelectedIndexChanged
        Try
            If String.IsNullOrWhiteSpace(txtPlaca1.Text) Then
                MsgBox(Me.Page, "Informe a Primeira Placa.")
                txtPlaca1.Focus()
                ddlEstadoPlaca1.SelectedIndex = 0
            Else
                txtMunicipioPlaca1.Text = String.Empty
                If Not String.IsNullOrWhiteSpace(ddlEstadoPlaca1.SelectedValue) Then
                    Session.Remove("objEstado1" & HID.Value)
                    Session("ssUF" & HID.Value) = ddlEstadoPlaca1.SelectedValue
                    Dim ucConsultaCodMunicipios = CType(Me.Page.FindControlRecursive("ucConsultaCodMunicipios"), ucConsultaCodMunicipios)
                    If ucConsultaCodMunicipios IsNot Nothing Then
                        ucConsultaCodMunicipios.SetarHID(HID.Value)
                        ucConsultaCodMunicipios.Limpar()
                    End If
                    Popup.ConsultaDeMunicipios(Me.Page, "objCidade1" & HID.Value)
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlEstadoPlaca2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlEstadoPlaca2.SelectedIndexChanged
        Try
            If String.IsNullOrWhiteSpace(txtPlaca2.Text) Then
                MsgBox(Me.Page, "Informe a Segunda Placa.")
                txtPlaca2.Focus()
                ddlEstadoPlaca2.SelectedIndex = 0
            Else
                txtMunicipioPlaca2.Text = String.Empty
                If Not String.IsNullOrWhiteSpace(ddlEstadoPlaca2.SelectedValue) Then
                    Session.Remove("objEstado2" & HID.Value)
                    Session("ssUF" & HID.Value) = ddlEstadoPlaca2.SelectedValue
                    Dim ucConsultaCodMunicipios = CType(Me.Page.FindControlRecursive("ucConsultaCodMunicipios"), ucConsultaCodMunicipios)
                    If ucConsultaCodMunicipios IsNot Nothing Then
                        ucConsultaCodMunicipios.SetarHID(HID.Value)
                        ucConsultaCodMunicipios.Limpar()
                    End If
                    Popup.ConsultaDeMunicipios(Me.Page, "objCidade2" & HID.Value)
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlEstadoPlaca3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlEstadoPlaca3.SelectedIndexChanged
        Try
            If String.IsNullOrWhiteSpace(txtPlaca3.Text) Then
                MsgBox(Me.Page, "Informe a Terceira Placa.")
                txtPlaca3.Focus()
                ddlEstadoPlaca3.SelectedIndex = 0
            Else
                txtMunicipioPlaca3.Text = String.Empty
                If Not String.IsNullOrWhiteSpace(ddlEstadoPlaca3.SelectedValue) Then
                    Session.Remove("objEstado3" & HID.Value)
                    Session("ssUF" & HID.Value) = ddlEstadoPlaca3.SelectedValue
                    Dim ucConsultaCodMunicipios = CType(Me.Page.FindControlRecursive("ucConsultaCodMunicipios"), ucConsultaCodMunicipios)
                    If ucConsultaCodMunicipios IsNot Nothing Then
                        ucConsultaCodMunicipios.SetarHID(HID.Value)
                        ucConsultaCodMunicipios.Limpar()
                    End If
                    Popup.ConsultaDeMunicipios(Me.Page, "objCidade3" & HID.Value)
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlEstadoPlaca4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlEstadoPlaca4.SelectedIndexChanged
        Try
            If String.IsNullOrWhiteSpace(txtPlaca4.Text) Then
                MsgBox(Me.Page, "Informe a Quarta Placa.")
                txtPlaca4.Focus()
                ddlEstadoPlaca4.SelectedIndex = 0
            Else
                txtMunicipioPlaca4.Text = String.Empty
                If Not String.IsNullOrWhiteSpace(ddlEstadoPlaca4.SelectedValue) Then
                    Session.Remove("objEstado4" & HID.Value)
                    Session("ssUF" & HID.Value) = ddlEstadoPlaca4.SelectedValue
                    Dim ucConsultaCodMunicipios = CType(Me.Page.FindControlRecursive("ucConsultaCodMunicipios"), ucConsultaCodMunicipios)
                    If ucConsultaCodMunicipios IsNot Nothing Then
                        ucConsultaCodMunicipios.SetarHID(HID.Value)
                        ucConsultaCodMunicipios.Limpar()
                    End If
                    Popup.ConsultaDeMunicipios(Me.Page, "objCidade4" & HID.Value)
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLiberar_Click(sender As Object, e As EventArgs) Handles lnkLiberar.Click
        radNao.Enabled = True
        radSim.Enabled = True
        txtObservacao.Enabled = True
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgLimparPlaca2_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles imgLimparPlaca2.Click
        Try
            LimparPlaca2()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgLimparPlaca3_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles imgLimparPlaca3.Click
        Try
            LimparPlaca3()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgLimparPlaca4_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles imgLimparPlaca4.Click
        Try
            LimparPlaca4()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgLimparMotorista_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles imgLimparMotorista.Click
        Try
            LimparMotorista()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnProprietario1_Click(sender As Object, e As EventArgs) Handles btnProprietario1.Click
        Try
            If String.IsNullOrWhiteSpace(txtPlaca1.Text) Then
                MsgBox(Me.Page, "A placa principal não foi informada.")
            Else
                ucConsultaClientes.Limpar()
                ucConsultaClientes.SetarTipoCliente("7,13")
                ucConsultaClientes.SetarTituloDIV("Consulta de Proprietário")
                Popup.ConsultaDeClientes(Me.Page, "objProprietario1" & HID.Value, "txtNome")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnProprietario2_Click(sender As Object, e As EventArgs) Handles btnProprietario2.Click
        Try
            If String.IsNullOrWhiteSpace(txtPlaca2.Text) Then
                MsgBox(Me.Page, "A segunda placa não foi informada.")
            Else
                ucConsultaClientes.Limpar()
                ucConsultaClientes.SetarTipoCliente("7,13")
                ucConsultaClientes.SetarTituloDIV("Consulta de Proprietário")
                Popup.ConsultaDeClientes(Me.Page, "objProprietario2" & HID.Value, "txtNome")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnProprietario3_Click(sender As Object, e As EventArgs) Handles btnProprietario3.Click
        Try
            If String.IsNullOrWhiteSpace(txtPlaca3.Text) Then
                MsgBox(Me.Page, "A terceira placa não foi informada.")
            Else
                ucConsultaClientes.Limpar()
                ucConsultaClientes.SetarTipoCliente("7,13")
                ucConsultaClientes.SetarTituloDIV("Consulta de Proprietário")
                Popup.ConsultaDeClientes(Me.Page, "objProprietario3" & HID.Value, "txtNome")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnProprietario4_Click(sender As Object, e As EventArgs) Handles btnProprietario4.Click
        Try
            If String.IsNullOrWhiteSpace(txtPlaca4.Text) Then
                MsgBox(Me.Page, "A quarta placa não foi informada.")
            Else
                ucConsultaClientes.Limpar()
                ucConsultaClientes.SetarTipoCliente("7,13")
                ucConsultaClientes.SetarTituloDIV("Consulta de Proprietário")
                Popup.ConsultaDeClientes(Me.Page, "objProprietario4" & HID.Value, "txtNome")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnMotorista_Click(sender As Object, e As EventArgs) Handles btnMotorista.Click
        Try
            ucConsultaClientes.Limpar()
            ucConsultaClientes.SetarTipoCliente("13")
            ucConsultaClientes.SetarTituloDIV("Consulta de Motorista")
            Popup.ConsultaDeClientes(Me.Page, "objMotorista" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnRepPro2_Click(sender As Object, e As EventArgs) Handles btnRepPro2.Click
        Try
            If String.IsNullOrWhiteSpace(txtPlaca2.Text) Then
                MsgBox(Me.Page, "Segunda Placa não foi informada.")
            ElseIf String.IsNullOrWhiteSpace(txtProprietario1.Text) Then
                MsgBox(Me.Page, "Proprietário da primeira placa não foi informado.")
            Else
                txtProprietario2.Text = txtProprietario1.Text
                txtCodigoProprietario2.Value = txtCodigoProprietario1.Value
                txtRNTRC2.Text = txtRNTRC1.Text
                ddlEstadoPlaca2.SelectedValue = ddlEstadoPlaca1.SelectedValue
                txtMunicipioPlaca2.Text = txtMunicipioPlaca1.Text
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnRepPro3_Click(sender As Object, e As EventArgs) Handles btnRepPro3.Click
        Try
            If String.IsNullOrWhiteSpace(txtPlaca3.Text) Then
                MsgBox(Me.Page, "Terceira Placa não foi informada.")
            ElseIf String.IsNullOrWhiteSpace(txtProprietario1.Text) Then
                MsgBox(Me.Page, "Proprietário da primeira placa não foi informado.")
            Else
                txtProprietario3.Text = txtProprietario1.Text
                txtCodigoProprietario3.Value = txtCodigoProprietario1.Value
                txtRNTRC3.Text = txtRNTRC1.Text
                ddlEstadoPlaca3.SelectedValue = ddlEstadoPlaca1.SelectedValue
                txtMunicipioPlaca3.Text = txtMunicipioPlaca1.Text
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnRepPro4_Click(sender As Object, e As EventArgs) Handles btnRepPro4.Click
        Try
            If String.IsNullOrWhiteSpace(txtPlaca4.Text) Then
                MsgBox(Me.Page, "Quarta Placa não foi informada.")
            ElseIf String.IsNullOrWhiteSpace(txtProprietario1.Text) Then
                MsgBox(Me.Page, "Proprietário da primeira placa não foi informado.")
            Else
                txtProprietario4.Text = txtProprietario1.Text
                txtCodigoProprietario4.Value = txtCodigoProprietario1.Value
                txtRNTRC4.Text = txtRNTRC1.Text
                ddlEstadoPlaca4.SelectedValue = ddlEstadoPlaca1.SelectedValue
                txtMunicipioPlaca4.Text = txtMunicipioPlaca1.Text
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If String.IsNullOrWhiteSpace(txtPlaca1.Text) Then
                MsgBox(Me.Page, "Informe o número da placa.")
                txtPlaca1.Focus()
                Exit Sub
            End If

            Dim objPlaca As New Placa()
            objPlaca.Placa01 = txtPlaca1.Text

            habilitarCampos(True)

            lnkConsultar.Parent.Visible = False

            If Not objPlaca.Selecionar Then
                MsgBox(Me.Page, "Placa não encontrada.")
                lnkNovo.Parent.Visible = True
            Else
                lnkAtualizar.Parent.Visible = True
                lnkExcluir.Parent.Visible = True

                If objPlaca.Proprietario01 IsNot Nothing Then
                    txtProprietario1.Text = Funcoes.FormatarListItemCliente(objPlaca.Proprietario01).Text
                    txtCodigoProprietario1.Value = Funcoes.FormatarListItemCliente(objPlaca.Proprietario01).Value
                    If objPlaca.EstadoPlaca01Detalhes IsNot Nothing Then
                        ddlEstadoPlaca1.SelectedValue = objPlaca.EstadoPlaca01Detalhes.Codigo.ToString()
                        txtMunicipioPlaca1.Text = objPlaca.CidadePlaca01.ToString()
                    End If
                    txtRNTRC1.Text = objPlaca.RNTRCPlaca01
                End If

                If Not String.IsNullOrWhiteSpace(objPlaca.Placa02) Then
                    txtPlaca2.Text = objPlaca.Placa02
                    If objPlaca.Proprietario02 IsNot Nothing AndAlso objPlaca.Proprietario01 IsNot Nothing Then
                        txtProprietario2.Text = Funcoes.FormatarListItemCliente(objPlaca.Proprietario02).Text
                        txtCodigoProprietario2.Value = Funcoes.FormatarListItemCliente(objPlaca.Proprietario02).Value
                    End If
                    If objPlaca.Estado02Detalhes IsNot Nothing Then
                        ddlEstadoPlaca2.SelectedValue = objPlaca.Estado02Detalhes.Codigo.ToString()
                        txtMunicipioPlaca2.Text = objPlaca.CidadePlaca02.ToString
                    End If
                    txtRNTRC2.Text = objPlaca.RNTRCPlaca02
                End If

                If Not String.IsNullOrWhiteSpace(objPlaca.Placa03) Then
                    txtPlaca3.Text = objPlaca.Placa03
                    If objPlaca.Proprietario03 IsNot Nothing AndAlso objPlaca.Proprietario02 IsNot Nothing Then
                        txtProprietario3.Text = Funcoes.FormatarListItemCliente(objPlaca.Proprietario03).Text
                        txtCodigoProprietario3.Value = Funcoes.FormatarListItemCliente(objPlaca.Proprietario03).Value
                    End If
                    If objPlaca.Estado03Detalhes IsNot Nothing Then
                        ddlEstadoPlaca3.SelectedValue = objPlaca.Estado03Detalhes.Codigo.ToString()
                        txtMunicipioPlaca3.Text = objPlaca.CidadePlaca03.ToString
                    End If
                    txtRNTRC3.Text = objPlaca.RNTRCPlaca03
                End If

                If Not String.IsNullOrWhiteSpace(objPlaca.Placa04) Then
                    txtPlaca4.Text = objPlaca.Placa04
                    If objPlaca.Proprietario04 IsNot Nothing AndAlso objPlaca.Proprietario03 IsNot Nothing Then
                        txtProprietario4.Text = Funcoes.FormatarListItemCliente(objPlaca.Proprietario04).Text
                        txtCodigoProprietario4.Value = Funcoes.FormatarListItemCliente(objPlaca.Proprietario04).Value
                    End If
                    If objPlaca.Estado04Detalhes IsNot Nothing Then
                        ddlEstadoPlaca4.SelectedValue = objPlaca.Estado04Detalhes.Codigo.ToString()
                        txtMunicipioPlaca4.Text = objPlaca.CidadePlaca04.ToString
                    End If
                    txtRNTRC4.Text = objPlaca.RNTRCPlaca04
                End If

                If Not objPlaca.TipoDeVeiculoDetalhes Is Nothing Then
                    ddlTipoDeVeiculo.SelectedValue = objPlaca.TipoDeVeiculoDetalhes.Codigo
                End If

                If Not objPlaca.ViaDeTransporteDetalhes Is Nothing Then
                    ddlViaDeTransporte.SelectedValue = objPlaca.ViaDeTransporteDetalhes.Codigo
                End If

                If Not String.IsNullOrWhiteSpace(objPlaca.CpfMotorista) AndAlso _
                    Not String.IsNullOrWhiteSpace(objPlaca.EndCpfMotorista) AndAlso _
                    Not objPlaca.Motorista Is Nothing AndAlso _
                    Not String.IsNullOrWhiteSpace(objPlaca.Motorista.Codigo) Then
                    Dim itemMotorista As ListItem = Funcoes.FormatarListItemCliente(objPlaca.Motorista)
                    With objPlaca.Motorista
                        txtMotorista.Text = .Nome
                        txtCodigoMotorista.Value = itemMotorista.Value
                        txtHabilitacao.Text = objPlaca.Habilitacao
                        txtEstadoMotorista.Text = .CodigoEstado
                        txtNomeEstadoMotorista.Text = .Estado.Descricao
                        txtMunicipioMotorista.Text = .Cidade
                    End With
                End If

                txtObservacao.Text = objPlaca.Observacao

                If objPlaca.Restricao = "S" Then
                    radSim.Checked = True
                Else
                    radNao.Checked = True
                End If
            End If

            If Funcoes.VerificaPermissao("CADASTRODEPLACAS", "LIBERAR") Then lnkLiberar.Parent.Visible = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("CADASTRODEPLACAS", "EXCLUIR") Then
                If String.IsNullOrWhiteSpace(txtPlaca1.Text) Then
                    MsgBox(Me.Page, "Informe a placa principal para efetuar exclusão.")
                Else
                    Dim objPlaca As New Placa(txtPlaca1.Text)

                    If objPlaca.TemLaudo Then
                        MsgBox(Me.Page, "Placa informada em Laudo não pode ser excluída.", eTitulo.Sucess)
                    ElseIf objPlaca.TemNota Then
                        MsgBox(Me.Page, "Placa informada em Nota Fiscal não pode ser excluída.", eTitulo.Sucess)
                    Else
                        objPlaca.IUD = "D"
                        If objPlaca.Salvar Then
                            MsgBox(Me.Page, "Placa Excluída com Sucesso.", eTitulo.Sucess)
                            Limpar()
                        End If
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("CADASTRODEPLACAS", "GRAVAR") Then
                If ValidaCampos() Then

                    Dim objPlacas As New Placa()
                    PreencheObjeto(objPlacas)
                    objPlacas.IUD = "I"

                    If objPlacas.Salvar Then
                        Limpar()
                        MsgBox(Me.Page, "Placa salva com Sucesso.", eTitulo.Sucess)
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("CADASTRODEPLACAS", "ALTERAR") Then
                If ValidaCampos() Then

                    Dim objPlacas As New Placa()
                    PreencheObjeto(objPlacas)
                    objPlacas.IUD = "U"

                    If objPlacas.Salvar Then
                        MsgBox(Me.Page, "Placa atualizada com Sucesso.", eTitulo.Sucess)
                        Limpar()
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "CadastroDePlacas")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

#End Region

#Region "Methods"

    Private Sub CarregarTipoVeiculo()
        ddlTipoDeVeiculo.Items.Clear()
        ddl.Carregar(ddlTipoDeVeiculo, CarregarDDL.Tabela.TipoDeVeiculo, "", True)
    End Sub

    Private Sub CarregarViaTransporte()
        ddlViaDeTransporte.Items.Clear()
        ddl.Carregar(ddlViaDeTransporte, CarregarDDL.Tabela.ViaDeTransportes, "", True)
    End Sub

    Private Sub Limpar()
        lnkConsultar.Parent.Visible = True
        lnkNovo.Parent.Visible = False
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
        lnkLiberar.Parent.Visible = False

        HID.Value = Guid.NewGuid().ToString()
        ucConsultaClientes.SetarHID(HID.Value)
        ucConsultaCodMunicipios.SetarHID(HID.Value)

        LimparPlaca1()
        LimparPlaca2()
        LimparPlaca3()
        LimparPlaca4()
        LimparMotorista()

        habilitarCampos(False)

        radNao.Enabled = False
        radSim.Enabled = False
        txtObservacao.Text = String.Empty
        txtObservacao.Enabled = False

        radNao.Checked = True

        txtPlaca1.Focus()
    End Sub

    Private Sub LimparPlaca1()
        txtPlaca1.Text = String.Empty
        txtRNTRC1.Text = String.Empty
        txtCodigoProprietario1.Value = String.Empty
        txtProprietario1.Text = String.Empty
        ddlEstadoPlaca1.SelectedIndex = 0
        txtMunicipioPlaca1.Text = String.Empty
    End Sub

    Private Sub LimparPlaca2()
        txtPlaca2.Text = String.Empty
        txtRNTRC2.Text = String.Empty
        txtCodigoProprietario2.Value = String.Empty
        txtProprietario2.Text = String.Empty
        ddlEstadoPlaca2.SelectedIndex = 0
        txtMunicipioPlaca2.Text = String.Empty
    End Sub

    Private Sub LimparPlaca3()
        txtPlaca3.Text = String.Empty
        txtRNTRC3.Text = String.Empty
        txtCodigoProprietario3.Value = String.Empty
        txtProprietario3.Text = String.Empty
        ddlEstadoPlaca3.SelectedIndex = 0
        txtMunicipioPlaca3.Text = String.Empty
    End Sub

    Private Sub LimparPlaca4()
        txtPlaca4.Text = String.Empty
        txtRNTRC4.Text = String.Empty
        txtCodigoProprietario4.Value = String.Empty
        txtProprietario4.Text = String.Empty
        ddlEstadoPlaca4.SelectedIndex = 0
        txtMunicipioPlaca4.Text = String.Empty
    End Sub

    Private Sub LimparMotorista()
        ddlTipoDeVeiculo.SelectedIndex = 0
        ddlViaDeTransporte.SelectedIndex = 0

        txtMotorista.Text = String.Empty
        txtCodigoMotorista.Value = String.Empty
        txtHabilitacao.Text = String.Empty
        txtEstadoMotorista.Text = String.Empty
        txtNomeEstadoMotorista.Text = String.Empty
        txtMunicipioMotorista.Text = String.Empty
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Session("objProprietario1" & HID.Value) IsNot Nothing Then
            Dim objCliente As Cliente = CType(obj, [Lib].Negocio.Cliente)
            txtProprietario1.Text = Funcoes.FormatarListItemCliente(objCliente).Text
            txtCodigoProprietario1.Value = Funcoes.FormatarListItemCliente(objCliente).Value
            txtRNTRC1.Text = objCliente.RNTRCTransportador
            If objCliente.Estado IsNot Nothing Then
                ddlEstadoPlaca1.SelectedValue = objCliente.CodigoEstado
                txtMunicipioPlaca1.Text = objCliente.Cidade
            End If
            Session.Remove("objProprietario1" & HID.Value)

        ElseIf Session("objProprietario2" & HID.Value) IsNot Nothing Then
            Dim objCliente As Cliente = CType(obj, [Lib].Negocio.Cliente)
            txtProprietario2.Text = Funcoes.FormatarListItemCliente(objCliente).Text
            txtCodigoProprietario2.Value = Funcoes.FormatarListItemCliente(objCliente).Value
            txtRNTRC2.Text = objCliente.RNTRCTransportador
            If objCliente.Estado IsNot Nothing Then
                ddlEstadoPlaca2.SelectedValue = objCliente.CodigoEstado
                txtMunicipioPlaca2.Text = objCliente.Cidade
            End If
            Session.Remove("objProprietario2" & HID.Value)

        ElseIf Session("objProprietario3" & HID.Value) IsNot Nothing Then
            Dim objCliente As Cliente = CType(obj, [Lib].Negocio.Cliente)
            txtProprietario3.Text = Funcoes.FormatarListItemCliente(objCliente).Text
            txtCodigoProprietario3.Value = Funcoes.FormatarListItemCliente(objCliente).Value
            txtRNTRC3.Text = objCliente.RNTRCTransportador
            If objCliente.Estado IsNot Nothing Then
                ddlEstadoPlaca3.SelectedValue = objCliente.CodigoEstado
                txtMunicipioPlaca3.Text = objCliente.Cidade
            End If
            Session.Remove("objProprietario3" & HID.Value)

        ElseIf Session("objProprietario4" & HID.Value) IsNot Nothing Then
            Dim objCliente As Cliente = CType(obj, [Lib].Negocio.Cliente)
            txtProprietario4.Text = Funcoes.FormatarListItemCliente(objCliente).Text
            txtCodigoProprietario4.Value = Funcoes.FormatarListItemCliente(objCliente).Value
            txtRNTRC4.Text = objCliente.RNTRCTransportador
            If objCliente.Estado IsNot Nothing Then
                ddlEstadoPlaca4.SelectedValue = objCliente.CodigoEstado
                txtMunicipioPlaca4.Text = objCliente.Cidade
            End If
            Session.Remove("objProprietario4" & HID.Value)

        ElseIf Session("objMotorista" & HID.Value) IsNot Nothing Then
            Dim objCliente As [Lib].Negocio.Cliente = CType(obj, [Lib].Negocio.Cliente)
            txtMotorista.Text = Funcoes.FormatarListItemCliente(objCliente).Text
            txtCodigoMotorista.Value = Funcoes.FormatarListItemCliente(objCliente).Value
            txtHabilitacao.Text = objCliente.Habilitacao
            txtEstadoMotorista.Text = objCliente.Estado.Codigo
            txtNomeEstadoMotorista.Text = objCliente.Estado.Descricao
            txtMunicipioMotorista.Text = objCliente.Cidade
            Session.Remove("objMotorista" & HID.Value)

        ElseIf Session("objEstado1" & HID.Value) IsNot Nothing Then
            Dim objEstado As [Lib].Negocio.Estado = obj
            Session.Remove("objEstado1" & HID.Value)

            Session("ssUF" & HID.Value) = objEstado.Codigo
            Dim ucConsultaCodMunicipios = CType(Me.Page.FindControlRecursive("ucConsultaCodMunicipios"), ucConsultaCodMunicipios)
            If ucConsultaCodMunicipios IsNot Nothing Then
                ucConsultaCodMunicipios.SetarHID(HID.Value)
                ucConsultaCodMunicipios.Limpar()
            End If
            Popup.ConsultaDeMunicipios(Me.Page, "objCidade1" & HID.Value)

        ElseIf Session("objCidade1" & HID.Value) IsNot Nothing Then
            Dim objMunicipio As [Lib].Negocio.Municipio = CType(Session("objCidade1" & HID.Value), [Lib].Negocio.Municipio)
            txtMunicipioPlaca1.Text = objMunicipio.CodigoMunicipio
            Session.Remove(HttpContext.Current.Session("ssTipoRetorno"))
            Session.Remove("objCidade1" & HID.Value)

        ElseIf Session("objEstado2" & HID.Value) IsNot Nothing Then
            Dim objEstado As [Lib].Negocio.Estado = obj
            Session.Remove("objEstado2" & HID.Value)

            Session("ssUF" & HID.Value) = objEstado.Codigo
            Dim ucConsultaCodMunicipios = CType(Me.Page.FindControlRecursive("ucConsultaCodMunicipios"), ucConsultaCodMunicipios)
            If ucConsultaCodMunicipios IsNot Nothing Then
                ucConsultaCodMunicipios.SetarHID(HID.Value)
                ucConsultaCodMunicipios.Limpar()
            End If
            Popup.ConsultaDeMunicipios(Me.Page, "objCidade2" & HID.Value)

        ElseIf Session("objCidade2" & HID.Value) IsNot Nothing Then
            Dim objMunicipio As [Lib].Negocio.Municipio = CType(Session("objCidade2" & HID.Value), [Lib].Negocio.Municipio)
            txtMunicipioPlaca2.Text = objMunicipio.CodigoMunicipio
            Session.Remove(HttpContext.Current.Session("ssTipoRetorno"))
            Session.Remove("objCidade2" & HID.Value)

        ElseIf Session("objEstado3" & HID.Value) IsNot Nothing Then
            Dim objEstado As [Lib].Negocio.Estado = obj
            Session.Remove("objEstado3" & HID.Value)

            Session("ssUF" & HID.Value) = objEstado.Codigo
            Dim ucConsultaCodMunicipios = CType(Me.Page.FindControlRecursive("ucConsultaCodMunicipios"), ucConsultaCodMunicipios)
            If ucConsultaCodMunicipios IsNot Nothing Then
                ucConsultaCodMunicipios.SetarHID(HID.Value)
                ucConsultaCodMunicipios.Limpar()
            End If
            Popup.ConsultaDeMunicipios(Me.Page, "objCidade3" & HID.Value)

        ElseIf Session("objCidade3" & HID.Value) IsNot Nothing Then
            Dim objMunicipio As [Lib].Negocio.Municipio = CType(Session("objCidade3" & HID.Value), [Lib].Negocio.Municipio)
            txtMunicipioPlaca3.Text = objMunicipio.CodigoMunicipio
            Session.Remove(HttpContext.Current.Session("ssTipoRetorno"))
            Session.Remove("objCidade3" & HID.Value)

        ElseIf Session("objEstado4" & HID.Value) IsNot Nothing Then
            Dim objEstado As [Lib].Negocio.Estado = obj
            Session.Remove("objEstado4" & HID.Value)

            Session("ssUF" & HID.Value) = objEstado.Codigo
            Dim ucConsultaCodMunicipios = CType(Me.Page.FindControlRecursive("ucConsultaCodMunicipios"), ucConsultaCodMunicipios)
            If ucConsultaCodMunicipios IsNot Nothing Then
                ucConsultaCodMunicipios.SetarHID(HID.Value)
                ucConsultaCodMunicipios.Limpar()
            End If
            Popup.ConsultaDeMunicipios(Me.Page, "objCidade4" & HID.Value)

        ElseIf Session("objCidade4" & HID.Value) IsNot Nothing Then
            Dim objMunicipio As [Lib].Negocio.Municipio = CType(Session("objCidade4" & HID.Value), [Lib].Negocio.Municipio)
            txtMunicipioPlaca4.Text = objMunicipio.CodigoMunicipio
            Session.Remove(HttpContext.Current.Session("ssTipoRetorno"))
            Session.Remove("objCidade4" & HID.Value)
        End If
    End Sub

    Private Sub PreencheObjeto(ByRef objPlaca As Placa)
        With objPlaca
            .Placa01 = txtPlaca1.Text.Trim()
            .RNTRCPlaca01 = txtRNTRC1.Text.Trim()

            If Not String.IsNullOrWhiteSpace(txtCodigoProprietario1.Value) Then
                .CodigoProprietario01 = txtCodigoProprietario1.Value.Split("-")(0)
                .EndProprietario01 = txtCodigoProprietario1.Value.Split("-")(1)
            End If
            .CidadePlaca01 = txtMunicipioPlaca1.Text.Trim()
            .EstadoPlaca01 = ddlEstadoPlaca1.SelectedValue

            'Placa 02
            If Not String.IsNullOrWhiteSpace(txtPlaca1.Text) Then
                .Placa02 = txtPlaca2.Text.Trim()
                .RNTRCPlaca02 = txtRNTRC2.Text.Trim()
                If Not String.IsNullOrWhiteSpace(txtCodigoProprietario2.Value) Then
                    .CodigoProprietario02 = txtCodigoProprietario2.Value.Split("-")(0)
                    .EndProprietario02 = txtCodigoProprietario2.Value.Split("-")(1)
                End If
                .CidadePlaca02 = txtMunicipioPlaca2.Text.Trim()
                .EstadoPlaca02 = ddlEstadoPlaca2.SelectedValue

            End If

            'Placa 03
            If Not String.IsNullOrWhiteSpace(txtPlaca2.Text) Then
                .Placa03 = txtPlaca3.Text.Trim()
                .RNTRCPlaca03 = txtRNTRC3.Text.Trim()
                If Not String.IsNullOrWhiteSpace(txtCodigoProprietario3.Value) Then
                    .CodigoProprietario03 = txtCodigoProprietario3.Value.Split("-")(0)
                    .EndProprietario03 = txtCodigoProprietario3.Value.Split("-")(1)
                End If
                .CidadePlaca03 = txtMunicipioPlaca3.Text.Trim()
                .EstadoPlaca03 = ddlEstadoPlaca3.SelectedValue
            End If

            'Placa 04
            If Not String.IsNullOrWhiteSpace(txtPlaca3.Text) Then
                .Placa04 = txtPlaca4.Text.Trim()
                .RNTRCPlaca04 = txtRNTRC4.Text.Trim()
                If Not String.IsNullOrWhiteSpace(txtCodigoProprietario4.Value) Then
                    .CodigoProprietario04 = txtCodigoProprietario4.Value.Split("-")(0)
                    .EndProprietario04 = txtCodigoProprietario4.Value.Split("-")(1)
                End If
                .CidadePlaca04 = txtMunicipioPlaca4.Text.Trim()
                .EstadoPlaca04 = ddlEstadoPlaca4.SelectedValue
            End If

            'Motorista
            If Not String.IsNullOrWhiteSpace(txtCodigoMotorista.Value) Then
                .CpfMotorista = txtCodigoMotorista.Value.Split("-")(0)
                .EndCpfMotorista = txtCodigoMotorista.Value.Split("-")(1)
            End If
            .Habilitacao = txtHabilitacao.Text.Trim()
            .TipoDeVeiculo = ddlTipoDeVeiculo.SelectedValue
            .ViaDeTransporte = ddlViaDeTransporte.SelectedValue
            .Restricao = IIf(radSim.Checked, "S", "N")
            .Observacao = txtObservacao.Text.Trim
        End With
    End Sub

    Function ValidarPlaca(Placa As String) As Boolean
        'For Each Str As String In Placa.Substring(0, 3)
        '    If (IsNumeric(Str)) Then
        '        Return False
        '    End If
        'Next
        'For Each Str As String In Placa.Substring(Placa.Length - 4, 4)
        '    If (Not IsNumeric(Str)) Then
        '        Return False
        '    End If
        'Next
        Return True
    End Function

    Function ValidaCampos() As Boolean
        If String.IsNullOrWhiteSpace(txtCodigoMotorista.Value) Then
            MsgBox(Me.Page, "Motorista não foi informado.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(ddlViaDeTransporte.SelectedValue) Then
            MsgBox(Me.Page, "Via de transporte não foi informada.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(ddlTipoDeVeiculo.SelectedValue) Then
            MsgBox(Me.Page, "Tipo de veículo não foi informado.")
            Return False
        ElseIf txtPlaca1.Text.Length > 0 And txtMunicipioPlaca1.Text.Length = 0 Then
            MsgBox(Me.Page, "Cidade da primeira placa não foi informada.")
            Return False
        ElseIf txtPlaca2.Text.Length > 0 And txtMunicipioPlaca2.Text.Length = 0 Then
            MsgBox(Me.Page, "Cidade da segunda placa não foi informada.")
            Return False
        ElseIf txtPlaca3.Text.Length > 0 And txtMunicipioPlaca3.Text.Length = 0 Then
            MsgBox(Me.Page, "Cidade da terceira placa não foi informada.")
            Return False
        ElseIf txtPlaca4.Text.Length > 0 And txtMunicipioPlaca4.Text.Length = 0 Then
            MsgBox(Me.Page, "Cidade da quarta placa não foi informada.")
            Return False
        ElseIf txtRNTRC1.Text.Length > 0 AndAlso txtCodigoProprietario1.Value.Length = 0 Then
            MsgBox(Me.Page, "Proprietário da primeira placa não foi informado.")
            Return False
        ElseIf txtCodigoProprietario1.Value.Length > 0 AndAlso txtRNTRC1.Text.Length = 0 Then
            MsgBox(Me.Page, "RNTRC da primeira placa não foi informado.")
            Return False
        ElseIf txtRNTRC2.Text.Length > 0 AndAlso txtCodigoProprietario2.Value.Length = 0 Then
            MsgBox(Me.Page, "Proprietário da segunda placa não foi informado.")
            Return False
        ElseIf txtCodigoProprietario2.Value.Length > 0 AndAlso txtRNTRC2.Text.Length = 0 Then
            MsgBox(Me.Page, "RNTRC da segunda placa não foi informado.")
            Return False
        ElseIf txtRNTRC3.Text.Length > 0 AndAlso txtCodigoProprietario3.Value.Length = 0 Then
            MsgBox(Me.Page, "Proprietário da terceira placa não foi informado.")
            Return False
        ElseIf txtCodigoProprietario3.Value.Length > 0 AndAlso txtRNTRC3.Text.Length = 0 Then
            MsgBox(Me.Page, "RNTRC da terceira placa não foi informado.")
            Return False
        ElseIf txtRNTRC4.Text.Length > 0 AndAlso txtCodigoProprietario4.Value.Length = 0 Then
            MsgBox(Me.Page, "Proprietário da quarta placa não foi informado.")
            Return False
        ElseIf txtCodigoProprietario4.Value.Length > 0 AndAlso txtRNTRC4.Text.Length = 0 Then
            MsgBox(Me.Page, "RNTRC da quarta placa não foi informado.")
            Return False
        ElseIf Not String.IsNullOrEmpty(txtPlaca1.Text) AndAlso String.IsNullOrEmpty(txtPlaca2.Text) AndAlso (Not String.IsNullOrEmpty(txtPlaca3.Text) OrElse Not String.IsNullOrEmpty(txtPlaca4.Text)) Then
            MsgBox(Me.Page, "A Placa 2 deve ser preenchida antes da placa 3 ou placa 4 .")
            Return False
        ElseIf Not String.IsNullOrEmpty(txtPlaca1.Text) AndAlso Not String.IsNullOrEmpty(txtPlaca2.Text) AndAlso String.IsNullOrEmpty(txtPlaca3.Text) AndAlso Not String.IsNullOrEmpty(txtPlaca4.Text) Then
            MsgBox(Me.Page, "A placa 2 e placa 3 devem ser preenchidas antes da placa 4.")
            Return False
        ElseIf radSim.Checked AndAlso String.IsNullOrWhiteSpace(txtObservacao.Text) Then
            MsgBox(Me.Page, "Indique o motivo pelo qual está restringindo a Placa!")
            txtObservacao.Focus()
            Return False
        ElseIf ddlViaDeTransporte.SelectedValue.Equals("1") AndAlso Not txtPlaca1.Text = "BONIFICA" Then 'Rodoviário
            If (Not String.IsNullOrEmpty(txtPlaca1.Text) AndAlso (Not ValidarPlaca(txtPlaca1.Text) OrElse txtPlaca1.Text.Trim().Replace("-", "").Length < 6)) Then
                MsgBox(Me.Page, "Placa 1 inválida.")
                Return False
            ElseIf (Not String.IsNullOrEmpty(txtPlaca2.Text) AndAlso (Not ValidarPlaca(txtPlaca2.Text) OrElse txtPlaca2.Text.Trim().Replace("-", "").Length < 6)) Then
                MsgBox(Me.Page, "Placa 2 inválida.")
                Return False
            ElseIf (Not String.IsNullOrEmpty(txtPlaca3.Text) AndAlso (Not ValidarPlaca(txtPlaca3.Text) OrElse txtPlaca3.Text.Trim().Replace("-", "").Length < 6)) Then
                MsgBox(Me.Page, "Placa 3 inválida.")
                Return False
            ElseIf (Not String.IsNullOrEmpty(txtPlaca4.Text) AndAlso (Not ValidarPlaca(txtPlaca4.Text) OrElse txtPlaca4.Text.Trim().Replace("-", "").Length < 6)) Then
                MsgBox(Me.Page, "Placa 4 inválida.")
                Return False
            End If
        End If
        Return True
    End Function

    Private Sub habilitarCampos(ByVal enabled As Boolean)
        txtPlaca1.Enabled = Not enabled
        txtRNTRC1.Enabled = enabled
        btnProprietario1.Enabled = enabled

        txtPlaca2.Enabled = enabled
        imgLimparPlaca2.Enabled = enabled
        txtRNTRC2.Enabled = enabled
        btnProprietario2.Enabled = enabled
        btnRepPro2.Enabled = enabled

        txtPlaca3.Enabled = enabled
        imgLimparPlaca3.Enabled = enabled
        txtRNTRC3.Enabled = enabled
        btnProprietario3.Enabled = enabled
        btnRepPro3.Enabled = enabled

        txtPlaca4.Enabled = enabled
        imgLimparPlaca4.Enabled = enabled
        txtRNTRC4.Enabled = enabled
        btnProprietario4.Enabled = enabled
        btnRepPro4.Enabled = enabled

        ddlTipoDeVeiculo.Enabled = enabled
        ddlViaDeTransporte.Enabled = enabled
        btnMotorista.Enabled = enabled
        imgLimparMotorista.Enabled = enabled
        txtHabilitacao.Enabled = enabled
    End Sub

#End Region

    Protected Sub radNao_CheckedChanged(sender As Object, e As EventArgs) Handles radNao.CheckedChanged
        txtObservacao.Text = String.Empty
        txtObservacao.Enabled = False
    End Sub

    Protected Sub radSim_CheckedChanged(sender As Object, e As EventArgs) Handles radSim.CheckedChanged
        txtObservacao.Enabled = True
        txtObservacao.Focus()
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("CADASTRODEPLACAS", "RELATORIO") Then
                Dim ds As New DataSet
                Dim Sql As String = ""

                Sql = "SELECT LTRIM(RTRIM(p1.Placa_Id)) as Placa_Id,                                                                      " & vbCrLf & _
                       "	vt.Descricao as ViaTransporte,                                                                                " & vbCrLf & _
                       "	tv.Descricao as TipoVeiculo,                                                                                  " & vbCrLf & _
                       "	p1.CidadePlaca as CidadePlaca,                                                                                " & vbCrLf & _
                       "	p1.EstadoPlaca as EstadoPlaca,                                                                                " & vbCrLf & _
                       "	isnull(p1.RNTRCPlaca,'') as RNTRCPlaca,                                                                       " & vbCrLf & _
                       "	p1.NomeMotorista as NomeMotorista,                                                                            " & vbCrLf & _
                       "	p1.CidadeMotorista as CidadeMotorista,                                                                        " & vbCrLf & _
                       "	p1.EstadoMotorista as EstadoMotorista,                                                                        " & vbCrLf & _
                       "	p1.Habilitacao as Habilitacao,                                                                                " & vbCrLf & _
                       "	p1.CpfMotorista as CpfMotorista,                                                                              " & vbCrLf & _
                       "	p1.EndCpfMotorista as EndCpfMotorista,                                                                        " & vbCrLf & _
                       "    isnull(p1.Proprietario,'') as Proprietario,                                                                   " & vbCrLf & _
                       "    isnull(p1.EndProprietario,'') as EndProprietario,                                                             " & vbCrLf & _
                       "    isnull(p1.NomeProprietario,'') as NomeProprietario,                                                           " & vbCrLf & _
                       "	LTRIM(RTRIM(p1.Placa01)) as Placa,	                                                                          " & vbCrLf & _
                       "	p1.CidadePlaca01 as CidadePlaca,                                                                              " & vbCrLf & _
                       "	p1.EstadoPlaca01 as EstadoPlaca,                                                                              " & vbCrLf & _
                       "	p1.RNTRCPlaca01 as RNTRCPlaca,                                                                                " & vbCrLf & _
                       "	p1.Proprietario01 as Proprietario,                                                                            " & vbCrLf & _
                       "	p1.EndProprietario01 as EndProprietario,                                                                      " & vbCrLf & _
                       "	p1.NomeProprietario01 as NomeProprietario                                                                     " & vbCrLf & _
                       "FROM Placas p1                                                                                                    " & vbCrLf & _
                       "INNER JOIN TiposDeVeiculos tv                                                                                     " & vbCrLf & _
                       "ON tv.Codigo_Id = p1.TipoVeiculo_Id                                                                               " & vbCrLf & _
                       "INNER JOIN ViaDeTransportes vt                                                                                    " & vbCrLf & _
                       "ON vt.Codigo_Id = p1.ViaTransporte_Id                                                                             " & vbCrLf & _
                       "WHERE 1=1                                                                                                         " & vbCrLf

                If Not String.IsNullOrWhiteSpace(txtPlaca1.Text) Then
                    Sql &= "AND p1.Placa_Id like '" & txtPlaca1.Text.Trim() & "%'" & vbCrLf
                End If

                If Not String.IsNullOrWhiteSpace(ddlViaDeTransporte.SelectedValue) Then
                    Sql &= "AND p1.ViaTransporte_Id = '" & ddlViaDeTransporte.SelectedValue & "' " & vbCrLf
                End If

                If Not String.IsNullOrWhiteSpace(ddlTipoDeVeiculo.SelectedValue) Then
                    Sql &= "AND p1.TipoVeiculo_Id = '" & ddlTipoDeVeiculo.SelectedValue & "' " & vbCrLf
                End If


                Sql &= "AND p1.Placa01 IS NOT NULL                                                                                                    " & vbCrLf & _
                        "AND p1.Placa01 <> ''                                                                                                         " & vbCrLf & _
                        "UNION ALL                                                                                                                    " & vbCrLf & _
                        "Select                                                                                                                       " & vbCrLf & _
                        "	LTRIM(RTRIM(p2.Placa_Id)) as Placa_Id,                                                                                    " & vbCrLf & _
                        "	vt.Descricao as ViaTransporte,                                                                                            " & vbCrLf & _
                        "	tv.Descricao as TipoVeiculo,                                                                                              " & vbCrLf & _
                        "	p2.CidadePlaca as CidadePlaca,                                                                                            " & vbCrLf & _
                        "	p2.EstadoPlaca as EstadoPlaca,                                                                                            " & vbCrLf & _
                        "	isnull(p2.RNTRCPlaca,'') as RNTRCPlaca,                                                                                   " & vbCrLf & _
                        "	p2.NomeMotorista as NomeMotorista,                                                                                        " & vbCrLf & _
                        "	p2.CidadeMotorista as CidadeMotorista,                                                                                    " & vbCrLf & _
                        "	p2.EstadoMotorista as EstadoMotorista,                                                                                    " & vbCrLf & _
                        "	p2.Habilitacao as Habilitacao,                                                                                            " & vbCrLf & _
                        "	p2.CpfMotorista as CpfMotorista,                                                                                          " & vbCrLf & _
                        "	p2.EndCpfMotorista as EndCpfMotorista,                                                                                    " & vbCrLf & _
                        "	isnull(p2.Proprietario,'') as Proprietario,                                                                               " & vbCrLf & _
                        "	isnull(p2.EndProprietario,'') as EndProprietario,                                                                         " & vbCrLf & _
                        "	isnull(p2.NomeProprietario,'') as NomeProprietario,                                                                       " & vbCrLf & _
                        "	LTRIM(RTRIM(p2.Placa02)) as Placa,                                                                                        " & vbCrLf & _
                        "	p2.CidadePlaca02 as CidadePlaca,                                                                                          " & vbCrLf & _
                        "	p2.EstadoPlaca02 as EstadoPlaca,                                                                                          " & vbCrLf & _
                        "	p2.RNTRCPlaca02 as RNTRCPlaca,                                                                                            " & vbCrLf & _
                        "	p2.Proprietario02 as Proprietario,                                                                                        " & vbCrLf & _
                        "	p2.EndProprietario02 as EndProprietario,                                                                                  " & vbCrLf & _
                        "	p2.NomeProprietario02 as NomeProprietario                                                                                 " & vbCrLf & _
                        "FROM Placas p2                                                                                                               " & vbCrLf & _
                        "INNER JOIN TiposDeVeiculos tv                                                                                                " & vbCrLf & _
                        "ON tv.Codigo_Id = p2.TipoVeiculo_Id                                                                                          " & vbCrLf & _
                        "INNER JOIN ViaDeTransportes vt                                                                                               " & vbCrLf & _
                        "ON vt.Codigo_Id = p2.ViaTransporte_Id                                                                                        " & vbCrLf & _
                        "WHERE 1=1                                                                                                                    " & vbCrLf

                If Not String.IsNullOrWhiteSpace(txtPlaca1.Text) Then
                    Sql &= "AND p2.Placa_Id like '" & txtPlaca1.Text.Trim() & "%' " & vbCrLf
                End If

                If Not String.IsNullOrWhiteSpace(ddlViaDeTransporte.SelectedValue) Then
                    Sql &= "AND p2.ViaTransporte_Id = '" & ddlViaDeTransporte.SelectedValue & "' " & vbCrLf
                End If

                If Not String.IsNullOrWhiteSpace(ddlTipoDeVeiculo.SelectedValue) Then
                    Sql &= "AND p2.TipoVeiculo_Id = '" & ddlTipoDeVeiculo.SelectedValue & "' " & vbCrLf
                End If

                Sql &= "AND p2.Placa02 IS NOT NULL                                                                                                    " & vbCrLf & _
                                   "AND p2.Placa02 <> ''                                                                                              " & vbCrLf & _
                                   "UNION ALL                                                                                                         " & vbCrLf & _
                                   "SELECT                                                                                                            " & vbCrLf & _
                                   "	LTRIM(RTRIM(p3.Placa_Id)) as Placa_Id,                                                                        " & vbCrLf & _
                                   "	vt.Descricao as ViaTransporte,                                                                                " & vbCrLf & _
                                   "	tv.Descricao as TipoVeiculo,                                                                                  " & vbCrLf & _
                                   "	p3.CidadePlaca as CidadePlaca,                                                                                " & vbCrLf & _
                                   "	p3.EstadoPlaca as EstadoPlaca,                                                                                " & vbCrLf & _
                                   "	isnull(p3.RNTRCPlaca,'') as RNTRCPlaca,                                                                       " & vbCrLf & _
                                   "	p3.NomeMotorista as NomeMotorista,                                                                            " & vbCrLf & _
                                   "	p3.CidadeMotorista as CidadeMotorista,                                                                        " & vbCrLf & _
                                   "	p3.EstadoMotorista as EstadoMotorista,                                                                        " & vbCrLf & _
                                   "	p3.Habilitacao as Habilitacao,                                                                                " & vbCrLf & _
                                   "	p3.CpfMotorista as CpfMotorista,                                                                              " & vbCrLf & _
                                   "	p3.EndCpfMotorista as EndCpfMotorista,                                                                        " & vbCrLf & _
                                   "	isnull(p3.Proprietario,'') as Proprietario,                                                                   " & vbCrLf & _
                                   "	isnull(p3.EndProprietario,'') as EndProprietario,                                                             " & vbCrLf & _
                                   "	isnull(p3.NomeProprietario,'') as NomeProprietario,                                                           " & vbCrLf & _
                                   "	LTRIM(RTRIM(p3.Placa03)) as Placa,                                                                            " & vbCrLf & _
                                   "	p3.CidadePlaca03 as CidadePlaca,                                                                              " & vbCrLf & _
                                   "	p3.EstadoPlaca03 as EstadoPlaca,                                                                              " & vbCrLf & _
                                   "	p3.RNTRCPlaca03 as RNTRCPlaca,                                                                                " & vbCrLf & _
                                   "	p3.Proprietario03 as Proprietario,                                                                            " & vbCrLf & _
                                   "	p3.EndProprietario03 as EndProprietario,                                                                      " & vbCrLf & _
                                   "	p3.NomeProprietario03 as NomeProprietario                                                                     " & vbCrLf & _
                                   "FROM Placas p3                                                                                                    " & vbCrLf & _
                                   "INNER JOIN TiposDeVeiculos tv                                                                                     " & vbCrLf & _
                                   "ON tv.Codigo_Id = p3.TipoVeiculo_Id                                                                               " & vbCrLf & _
                                   "INNER JOIN ViaDeTransportes vt                                                                                    " & vbCrLf & _
                                   "ON vt.Codigo_Id = p3.ViaTransporte_Id                                                                             " & vbCrLf & _
                                   "WHERE 1=1                                                                                                         " & vbCrLf

                If Not String.IsNullOrWhiteSpace(txtPlaca1.Text) Then
                    Sql &= "AND p3.Placa_Id like '" & txtPlaca1.Text.Trim() & "%' " & vbCrLf
                End If

                If Not String.IsNullOrWhiteSpace(ddlViaDeTransporte.SelectedValue) Then
                    Sql &= "AND p3.ViaTransporte_Id = '" & ddlViaDeTransporte.SelectedValue & "' " & vbCrLf
                End If

                If Not String.IsNullOrWhiteSpace(ddlTipoDeVeiculo.SelectedValue) Then
                    Sql &= "AND p3.TipoVeiculo_Id = '" & ddlTipoDeVeiculo.SelectedValue & "' " & vbCrLf
                End If

                Sql &= "AND p3.Placa03 IS NOT NULL                                                                                        " & vbCrLf & _
                       "AND p3.Placa03 <> ''                                                                                              " & vbCrLf

                ds = Banco.ConsultaDataSet(Sql, "Placas")

                Dim parameters As New Dictionary(Of String, Object)
                parameters.Add("ConsultaParametros", getParam())

                Funcoes.BindReport(Me.Page, ds, "Cr_Placas", eExportType.PDF, parameters)

            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function getParam() As String
        Dim param As String = "Parametros:" & vbCrLf
        If Not String.IsNullOrWhiteSpace(txtPlaca1.Text) Then
            param &= "Placa: " & txtPlaca1.Text & " - "
        End If
        If Not String.IsNullOrWhiteSpace(txtProprietario1.Text) Then
            param &= "Proprietário: " & txtProprietario1.Text & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(ddlViaDeTransporte.SelectedValue) Then
            param &= "Via de Transporte: " & ddlViaDeTransporte.SelectedValue & " - "
        End If
        If Not String.IsNullOrWhiteSpace(ddlTipoDeVeiculo.SelectedValue) Then
            param &= "Tipo de Veículo: " & ddlTipoDeVeiculo.SelectedValue & "."
        End If

        Return param
    End Function


End Class