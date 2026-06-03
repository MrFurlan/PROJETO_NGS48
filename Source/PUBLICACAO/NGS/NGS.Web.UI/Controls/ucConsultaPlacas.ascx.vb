Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucConsultaPlacas
    Inherits BaseUserControl

    Dim objPlaca As [Lib].Negocio.Placa

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not IsPostBack Then
                ddl.Carregar(ddlTipoDeVeiculo, CarregarDDL.Tabela.TipoDeVeiculo, "", True)
                ddl.Carregar(ddlViaDeTransporte, CarregarDDL.Tabela.ViaDeTransportes, "", True)
                CarregarEstados()
                txtPlaca1.Focus()
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

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
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
                    Session.Remove("objEstadoPXE1" & HID.Value)
                    Session("ssUF" & HID.Value) = ddlEstadoPlaca1.SelectedValue
                    Dim ucConsultaCodMunicipios = CType(Me.Page.FindControlRecursive("ucConsultaCodMunicipios"), ucConsultaCodMunicipios)
                    If ucConsultaCodMunicipios IsNot Nothing Then
                        ucConsultaCodMunicipios.SetarHID(HID.Value)
                        ucConsultaCodMunicipios.Limpar()
                        ucConsultaCodMunicipios.MainUserControl = Me
                    End If
                    Popup.ConsultaDeMunicipios(Me.Page, "objCidadePXC1" & HID.Value)
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
                    Session.Remove("objEstadoPXE2" & HID.Value)
                    Session("ssUF" & HID.Value) = ddlEstadoPlaca2.SelectedValue
                    Dim ucConsultaCodMunicipios = CType(Me.Page.FindControlRecursive("ucConsultaCodMunicipios"), ucConsultaCodMunicipios)
                    If ucConsultaCodMunicipios IsNot Nothing Then
                        ucConsultaCodMunicipios.Limpar()
                        ucConsultaCodMunicipios.MainUserControl = Me
                    End If
                    Popup.ConsultaDeMunicipios(Me.Page, "objCidadePXC2" & HID.Value)
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
                    Session.Remove("objEstadoPXE3" & HID.Value)
                    Session("ssUF" & HID.Value) = ddlEstadoPlaca3.SelectedValue
                    Dim ucConsultaCodMunicipios = CType(Me.Page.FindControlRecursive("ucConsultaCodMunicipios"), ucConsultaCodMunicipios)
                    If ucConsultaCodMunicipios IsNot Nothing Then
                        ucConsultaCodMunicipios.Limpar()
                        ucConsultaCodMunicipios.MainUserControl = Me
                    End If
                    Popup.ConsultaDeMunicipios(Me.Page, "objCidadePXC3" & HID.Value)
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
                    Session.Remove("objEstadoPXE4" & HID.Value)
                    Session("ssUF" & HID.Value) = ddlEstadoPlaca4.SelectedValue
                    Dim ucConsultaCodMunicipios = CType(Me.Page.FindControlRecursive("ucConsultaCodMunicipios"), ucConsultaCodMunicipios)
                    If ucConsultaCodMunicipios IsNot Nothing Then
                        ucConsultaCodMunicipios.Limpar()
                        ucConsultaCodMunicipios.MainUserControl = Me
                    End If
                    Popup.ConsultaDeMunicipios(Me.Page, "objCidadePXC4" & HID.Value)
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Session("objProprietarioPLXP1" & HID.Value) IsNot Nothing Then
            Dim objCliente As [Lib].Negocio.Cliente = CType(obj, [Lib].Negocio.Cliente)
            txtProprietario1.Text = Funcoes.FormatarListItemCliente(objCliente).Text
            txtCodigoProprietario1.Value = Funcoes.FormatarListItemCliente(objCliente).Value
            txtRNTRC1.Text = objCliente.RNTRCTransportador
            txtProprietario1.Enabled = False

            If String.IsNullOrWhiteSpace(txtMunicipioPlaca1.Text) Then
                ddlEstadoPlaca1.SelectedValue = objCliente.CodigoEstado
                txtMunicipioPlaca1.Text = objCliente.Cidade
            End If

            Session.Remove("objProprietarioPLXP1" & HID.Value)
        ElseIf Session("objProprietarioPLXP2" & HID.Value) IsNot Nothing Then
            Dim objCliente As [Lib].Negocio.Cliente = CType(obj, [Lib].Negocio.Cliente)

            txtProprietario2.Text = Funcoes.FormatarListItemCliente(objCliente).Text
            txtCodigoProprietario2.Value = Funcoes.FormatarListItemCliente(objCliente).Value

            txtRNTRC2.Text = objCliente.RNTRCTransportador
            txtProprietario2.Enabled = False

            If String.IsNullOrWhiteSpace(txtMunicipioPlaca2.Text) Then
                ddlEstadoPlaca2.SelectedValue = objCliente.CodigoEstado
                txtMunicipioPlaca2.Text = objCliente.Cidade
            End If
            Session.Remove("objProprietarioPLXP2" & HID.Value)
        ElseIf Session("objProprietarioPLXP3" & HID.Value) IsNot Nothing Then
            Dim objCliente As [Lib].Negocio.Cliente = CType(obj, [Lib].Negocio.Cliente)
            txtProprietario3.Text = Funcoes.FormatarListItemCliente(objCliente).Text
            txtCodigoProprietario3.Value = Funcoes.FormatarListItemCliente(objCliente).Value

            txtRNTRC3.Text = objCliente.RNTRCTransportador
            txtProprietario3.Enabled = False

            If String.IsNullOrWhiteSpace(txtMunicipioPlaca3.Text) Then
                ddlEstadoPlaca3.SelectedValue = objCliente.CodigoEstado
                txtMunicipioPlaca3.Text = objCliente.Cidade
            End If

            Session.Remove("objProprietarioPLXP3" & HID.Value)
        ElseIf Session("objProprietarioPLXP4" & HID.Value) IsNot Nothing Then
            Dim objCliente As [Lib].Negocio.Cliente = CType(obj, [Lib].Negocio.Cliente)

            txtProprietario1.Text = Funcoes.FormatarListItemCliente(objCliente).Text
            txtCodigoProprietario1.Value = Funcoes.FormatarListItemCliente(objCliente).Value
            txtRNTRC4.Text = objCliente.RNTRCTransportador
            txtProprietario4.Enabled = False

            If String.IsNullOrWhiteSpace(txtMunicipioPlaca4.Text) Then
                ddlEstadoPlaca4.SelectedValue = objCliente.CodigoEstado
                txtMunicipioPlaca4.Text = objCliente.Cidade
            End If

            Session.Remove("objProprietarioPLXP4" & HID.Value)
        ElseIf Session("objMotoristaPXE" & HID.Value) IsNot Nothing Then
            Dim objCliente As [Lib].Negocio.Cliente = CType(obj, [Lib].Negocio.Cliente)
            txtMotorista.Text = objCliente.Nome
            txtCodigoMotorista.Value = objCliente.Codigo & "-" & objCliente.CodigoEndereco
            txtEstadoMotorista.Text = objCliente.CodigoEstado
            txtNomeEstadoMotorista.Text = objCliente.Estado.Descricao
            txtMunicipioMotorista.Text = objCliente.Cidade
            txtHabilitacao.Text = objCliente.Habilitacao
            txtEstadoMotorista.Text = objCliente.Estado.Codigo
            txtNomeEstadoMotorista.Text = objCliente.Estado.Descricao
            txtMunicipioMotorista.Text = objCliente.Cidade
            Session.Remove("objMotoristaPXE" & HID.Value)
        ElseIf Session("objEstadoPXE1" & HID.Value) IsNot Nothing Then
            Dim objEstado As [Lib].Negocio.Estado = obj
            Session.Remove("objEstadoPXE1" & HID.Value)
            Session("ssUF" & HID.Value) = objEstado.Codigo
            Dim ucConsultaCodMunicipios = CType(Me.Page.FindControlRecursive("ucConsultaCodMunicipios"), ucConsultaCodMunicipios)
            If ucConsultaCodMunicipios IsNot Nothing Then
                ucConsultaCodMunicipios.SetarHID(HID.Value)
                ucConsultaCodMunicipios.Limpar()
                ucConsultaCodMunicipios.MainUserControl = Me
            End If
            Popup.ConsultaDeMunicipios(Me.Page, "objCidadePXC1" & HID.Value)
        ElseIf Session("objEstadoPXE2" & HID.Value) IsNot Nothing Then
            Dim objEstado As [Lib].Negocio.Estado = obj
            Session.Remove("objEstadoPXE2" & HID.Value)
            Session("ssUF" & HID.Value) = objEstado.Codigo
            Dim ucConsultaCodMunicipios = CType(Me.Page.FindControlRecursive("ucConsultaCodMunicipios"), ucConsultaCodMunicipios)
            If ucConsultaCodMunicipios IsNot Nothing Then
                ucConsultaCodMunicipios.Limpar()
                ucConsultaCodMunicipios.MainUserControl = Me
            End If
            Popup.ConsultaDeMunicipios(Me.Page, "objCidadePXC2" & HID.Value)
        ElseIf Session("objEstadoPXE3" & HID.Value) IsNot Nothing Then
            Dim objEstado As [Lib].Negocio.Estado = obj
            Session.Remove("objEstadoPXE3" & HID.Value)
            Session("ssUF" & HID.Value) = objEstado.Codigo
            Dim ucConsultaCodMunicipios = CType(Me.Page.FindControlRecursive("ucConsultaCodMunicipios"), ucConsultaCodMunicipios)
            If ucConsultaCodMunicipios IsNot Nothing Then
                ucConsultaCodMunicipios.Limpar()
                ucConsultaCodMunicipios.MainUserControl = Me
            End If
            Popup.ConsultaDeMunicipios(Me.Page, "objCidadePXC3" & HID.Value)
        ElseIf Session("objEstadoPXE4" & HID.Value) IsNot Nothing Then
            Dim objEstado As [Lib].Negocio.Estado = obj
            Session.Remove("objEstadoPXE4" & HID.Value)
            Session("ssUF" & HID.Value) = objEstado.Codigo
            Dim ucConsultaCodMunicipios = CType(Me.Page.FindControlRecursive("ucConsultaCodMunicipios"), ucConsultaCodMunicipios)
            If ucConsultaCodMunicipios IsNot Nothing Then
                ucConsultaCodMunicipios.Limpar()
                ucConsultaCodMunicipios.MainUserControl = Me
            End If
            Popup.ConsultaDeMunicipios(Me.Page, "objCidadePXC4" & HID.Value)
        ElseIf Session("objCidadePXC1" & HID.Value) IsNot Nothing Then
            Dim objMunicipio As [Lib].Negocio.Municipio = CType(Session("objCidadePXC1" & HID.Value), [Lib].Negocio.Municipio)
            txtMunicipioPlaca1.Text = objMunicipio.CodigoMunicipio
            Session.Remove(Session("ssTipoRetorno"))
            Session.Remove("objCidadePXC1" & HID.Value)
        ElseIf Session("objCidadePXC2" & HID.Value) IsNot Nothing Then
            Dim objMunicipio As [Lib].Negocio.Municipio = CType(Session("objCidadePXC2" & HID.Value), [Lib].Negocio.Municipio)
            txtMunicipioPlaca2.Text = objMunicipio.CodigoMunicipio
            Session.Remove(Session("ssTipoRetorno"))
            Session.Remove("objCidadePXC2" & HID.Value)
        ElseIf Session("objCidadePXC3" & HID.Value) IsNot Nothing Then
            Dim objMunicipio As [Lib].Negocio.Municipio = CType(Session("objCidadePXC3" & HID.Value), [Lib].Negocio.Municipio)
            txtMunicipioPlaca3.Text = objMunicipio.CodigoMunicipio
            Session.Remove(Session("ssTipoRetorno"))
            Session.Remove("objCidadePXC3" & HID.Value)
        ElseIf Session("objCidadePXC4" & HID.Value) IsNot Nothing Then
            Dim objMunicipio As [Lib].Negocio.Municipio = CType(Session("objCidadePXC4" & HID.Value), [Lib].Negocio.Municipio)
            txtMunicipioPlaca4.Text = objMunicipio.CodigoMunicipio
            Session.Remove(Session("ssTipoRetorno"))
            Session.Remove("objCidadePXC4" & HID.Value)
        End If
    End Sub

    Public Overrides Sub Limpar()
        lnkConsultar.Parent.Visible = True
        lnkConfirmar.Parent.Visible = False

        DirectCast(Me.NamingContainer.FindControl("ucConsultaClientes"), ucConsultaClientes).SetarHID(HID.Value)
        DirectCast(Me.NamingContainer.FindControl("ucConsultaCodMunicipios"), ucConsultaCodMunicipios).SetarHID(HID.Value)

        LimparPlaca1()

        LimparPlaca2()
        LimparPlaca3()
        LimparPlaca4()

        LimparMotorista()

        Bloquear(False)

        txtObservacao.Text = String.Empty
        txtObservacao.Enabled = False
        radNao.Checked = True

        Session.Remove("ssMessage")
        Session.Remove("_MainUserControl")
        Session.Remove("objPlaca" & HID.Value)
        txtPlaca1.Focus()
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

    Private Sub CarregarVariavelPlaca()
        objPlaca = CType(Session("objPlaca" & HID.Value.ToString), [Lib].Negocio.Placa)
        'Session("_MainUserControl") = Nothing
        If (objPlaca Is Nothing) Then
            objPlaca = New [Lib].Negocio.Placa
            objPlaca.IUD = "I"
        End If
    End Sub

    Private Sub SalvarVariavelPlaca()
        Session("objPlaca" & HID.Value.ToString) = objPlaca
    End Sub

    Private Sub Confirmar()
        If Not ValidaCampos() Then
            Exit Sub
        End If

        CarregarVariavelPlaca()
        objPlaca = New [Lib].Negocio.Placa(txtPlaca1.Text)

        Dim Placa As String = String.Empty

        If Left(Session("ssEmpresa"), 8) = "05272759" Then
            Placa = txtPlaca1.Text
        Else
            Placa = Funcoes.EliminarCaracteresEspeciais(txtPlaca1.Text)
        End If

        Placa = Placa.Replace(".", "").Replace(",", "").Replace(" ", "")

        objPlaca.Placa01 = Trim(Placa).ToUpper()
        objPlaca.ViaDeTransporte = ddlViaDeTransporte.SelectedValue
        objPlaca.TipoDeVeiculo = ddlTipoDeVeiculo.SelectedValue
        objPlaca.Restricao = IIf(radSim.Checked, "S", "N")
        objPlaca.Observacao = txtObservacao.Text.Trim

        objPlaca.CodigoProprietario02 = ""
        objPlaca.EndProprietario02 = 0

        objPlaca.CodigoProprietario03 = ""
        objPlaca.EndProprietario03 = 0

        objPlaca.CodigoProprietario04 = ""
        objPlaca.EndProprietario04 = 0

        If txtCodigoProprietario1.Value.Length > 0 Then
            Dim Proprietario() As String = txtCodigoProprietario1.Value.ToString.Split("-")
            objPlaca.CodigoProprietario01 = Proprietario(0)
            objPlaca.EndProprietario01 = Proprietario(1)
        End If

        If Trim(txtPlaca2.Text).Length = 0 Then
            txtPlaca2.Text = ""
            txtMunicipioPlaca2.Text = ""
            txtRNTRC2.Text = ""
            txtCodigoProprietario2.Value = ""
        Else
            If txtProprietario2.Text.Length > 0 Then
                If txtCodigoProprietario2.Value.Length > 0 Then
                    Dim Proprietario() As String = txtCodigoProprietario2.Value.ToString.Split("-")
                    objPlaca.CodigoProprietario02 = Proprietario(0)
                    objPlaca.EndProprietario02 = Proprietario(1)
                Else
                    objPlaca.CodigoProprietario02 = Funcoes.EliminarCaracteresEspeciais(txtProprietario2.Text).Replace("-", "").Replace("/", "").Replace(".", "")
                    objPlaca.EndProprietario02 = 0

                    If objPlaca.CodigoProprietario02.Length = 11 Then
                        If Not Funcoes.ValidaCPF(objPlaca.CodigoProprietario02) Then
                            MsgBox(Me.Page, "CPF do Proprietário da Segunda Placa inválido.", eTitulo.Info)
                            Exit Sub
                        End If
                    Else
                        If Not Funcoes.ValidaCNPJ(objPlaca.CodigoProprietario02) Then
                            MsgBox(Me.Page, "CNPJ do Proprietário da Segunda Placa inválido.", eTitulo.Info)
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If

        If Trim(txtPlaca3.Text).Length = 0 Then
            txtPlaca3.Text = ""
            txtMunicipioPlaca3.Text = ""
            txtRNTRC3.Text = ""
            txtCodigoProprietario3.Value = ""
        Else
            If txtProprietario3.Text.Length > 0 Then
                If txtCodigoProprietario3.Value.Length > 0 Then
                    Dim Proprietario() As String = txtCodigoProprietario3.Value.ToString.Split("-")
                    objPlaca.CodigoProprietario03 = Proprietario(0)
                    objPlaca.EndProprietario03 = Proprietario(1)
                Else
                    objPlaca.CodigoProprietario03 = Funcoes.EliminarCaracteresEspeciais(txtProprietario3.Text).Replace("-", "").Replace("/", "").Replace(".", "")
                    objPlaca.EndProprietario03 = 0

                    If objPlaca.CodigoProprietario03.Length = 11 Then
                        If Not Funcoes.ValidaCPF(objPlaca.CodigoProprietario03) Then
                            MsgBox(Me.Page, "CPF do Proprietário da Terceira Placa inválido.")
                            Exit Sub
                        End If
                    Else
                        If Not Funcoes.ValidaCNPJ(objPlaca.CodigoProprietario03) Then
                            MsgBox(Me.Page, "CNPJ do Proprietário da Terceira Placa inválido.")
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If

        If Trim(txtPlaca4.Text).Length = 0 Then
            txtPlaca4.Text = ""
            txtMunicipioPlaca4.Text = ""
            txtRNTRC4.Text = ""
            txtCodigoProprietario4.Value = ""
        Else
            If txtProprietario4.Text.Length > 0 Then
                If txtCodigoProprietario4.Value.Length > 0 Then
                    Dim Proprietario() As String = txtCodigoProprietario4.Value.ToString.Split("-")
                    objPlaca.CodigoProprietario04 = Proprietario(0)
                    objPlaca.EndProprietario04 = Proprietario(1)
                Else
                    objPlaca.CodigoProprietario04 = Funcoes.EliminarCaracteresEspeciais(txtProprietario4.Text).Replace("-", "").Replace("/", "").Replace(".", "")
                    objPlaca.EndProprietario04 = 0

                    If objPlaca.CodigoProprietario04.Length = 11 Then
                        If Not Funcoes.ValidaCPF(objPlaca.CodigoProprietario04) Then
                            MsgBox(Me.Page, "CPF do Proprietário da Quarta Placa inválido.")
                            Exit Sub
                        End If
                    Else
                        If Not Funcoes.ValidaCNPJ(objPlaca.CodigoProprietario04) Then
                            MsgBox(Me.Page, "CNPJ do Proprietário da Quarta Placa inválido.")
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If

        objPlaca.Placa02 = Trim(txtPlaca2.Text).ToUpper()
        objPlaca.Placa03 = Trim(txtPlaca3.Text).ToUpper()
        objPlaca.Placa04 = Trim(txtPlaca4.Text).ToUpper()
        objPlaca.RNTRCPlaca01 = Trim(txtRNTRC1.Text)
        objPlaca.RNTRCPlaca02 = Trim(txtRNTRC2.Text)
        objPlaca.RNTRCPlaca03 = Trim(txtRNTRC3.Text)
        objPlaca.RNTRCPlaca04 = Trim(txtRNTRC4.Text)

        Dim Motorista() As String = txtCodigoMotorista.Value.Split("-")
        objPlaca.CpfMotorista = Funcoes.EliminarCaracteresEspeciais(Motorista(0))
        objPlaca.EndCpfMotorista = Motorista(1)
        Dim objMotorista = New [Lib].Negocio.Cliente(objPlaca.CpfMotorista, objPlaca.EndCpfMotorista)
        objPlaca.Motorista = objMotorista

        objPlaca.Habilitacao = Funcoes.EliminarCaracteresEspeciais(txtHabilitacao.Text)
        objPlaca.CidadePlaca01 = txtMunicipioPlaca1.Text.Trim
        objPlaca.CidadePlaca02 = txtMunicipioPlaca2.Text.Trim
        objPlaca.CidadePlaca03 = txtMunicipioPlaca3.Text.Trim
        objPlaca.CidadePlaca04 = txtMunicipioPlaca4.Text.Trim

        objPlaca.EstadoPlaca01 = ddlEstadoPlaca1.SelectedValue
        objPlaca.EstadoPlaca02 = ddlEstadoPlaca2.SelectedValue
        objPlaca.EstadoPlaca03 = ddlEstadoPlaca3.SelectedValue
        objPlaca.EstadoPlaca04 = ddlEstadoPlaca4.SelectedValue



        If objPlaca.Salvar Then
            Session(Session("ssTipoRetorno")) = objPlaca
            If MainUserControl IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(MainUserControl.ClientID) Then
                Dim ucName = MainUserControl.ClientID.Split("_")
                Dim uc As UserControl = CType(WebHelpers.FindControlRecursive(Me.Page, ucName(ucName.Length - 1)), UserControl)
                CType(uc, IBaseUserControl).Carregar(objPlaca)
            Else
                Session("objPlaca" & HID.Value) = objPlaca
                CType(Me.Page, IBasePage).Carregar(objPlaca)
            End If
            Popup.CloseDialog(Me.Page, "divConsultaPlacas")
        Else
            MsgBox(Me.Page, Session("ssMessage"))
        End If
    End Sub

    Protected Sub imgLimparMotorista_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imgLimparMotorista.Click
        LimparMotorista()
    End Sub

    Function ValidaCampos() As Boolean
        If txtCodigoMotorista.Value.Length = 0 Then
            MsgBox(Me.Page, "Motorista não foi informado.", eTitulo.Info)
            Return False
        ElseIf txtMotorista.Text.Length = 0 Then
            MsgBox(Me.Page, "Nome do Motorista não foi informado.", eTitulo.Info)
            Return False
        ElseIf String.IsNullOrWhiteSpace(ddlViaDeTransporte.SelectedValue) Then
            MsgBox(Me.Page, "Via de transporte não foi informada.", eTitulo.Info)
            Return False
        ElseIf String.IsNullOrWhiteSpace(ddlTipoDeVeiculo.SelectedValue) Then
            MsgBox(Me.Page, "Tipo de veículo não foi informado.", eTitulo.Info)
            Return False
        ElseIf txtPlaca1.Text.Length > 0 And txtMunicipioPlaca1.Text.Length = 0 Then
            MsgBox(Me.Page, "Cidade da primeira placa não foi informada.", eTitulo.Info)
            Return False
        ElseIf txtPlaca2.Text.Length > 0 And txtMunicipioPlaca2.Text.Length = 0 Then
            MsgBox(Me.Page, "Cidade da segunda placa não foi informada.", eTitulo.Info)
            Return False
        ElseIf txtPlaca3.Text.Length > 0 And txtMunicipioPlaca3.Text.Length = 0 Then
            MsgBox(Me.Page, "Cidade da terceira placa não foi informada.", eTitulo.Info)
            Return False
        ElseIf txtPlaca4.Text.Length > 0 And txtMunicipioPlaca4.Text.Length = 0 Then
            MsgBox(Me.Page, "Cidade da quarta placa não foi informada.", eTitulo.Info)
            Return False
        ElseIf txtRNTRC1.Text.Length > 0 AndAlso txtCodigoProprietario1.Value.Length = 0 Then
            MsgBox(Me.Page, "Proprietário da Primeira Placa não foi informado.")
            Return False
        ElseIf txtCodigoProprietario1.Value.Length > 0 AndAlso txtRNTRC1.Text.Length = 0 Then
            MsgBox(Me.Page, "RNTRC da Primeira Placa não foi informado.")
            Return False
        ElseIf txtRNTRC2.Text.Length > 0 AndAlso txtCodigoProprietario2.Value.Length = 0 Then
            MsgBox(Me.Page, "Proprietário da Segunda Placa não foi informado.")
            Return False
        ElseIf txtCodigoProprietario2.Value.Length > 0 AndAlso txtRNTRC2.Text.Length = 0 Then
            MsgBox(Me.Page, "RNTRC da Segunda Placa não foi informado.")
            Return False
        ElseIf txtRNTRC3.Text.Length > 0 AndAlso txtCodigoProprietario3.Value.Length = 0 Then
            MsgBox(Me.Page, "Proprietário da Terceira Placa não foi informado.")
            Return False
        ElseIf txtCodigoProprietario3.Value.Length > 0 AndAlso txtRNTRC3.Text.Length = 0 Then
            MsgBox(Me.Page, "RNTRC da Terceira Placa não foi informado.")
            Return False
        ElseIf txtRNTRC4.Text.Length > 0 AndAlso txtCodigoProprietario4.Value.Length = 0 Then
            MsgBox(Me.Page, "Proprietário da Quarta Placa não foi informado.")
            Return False
        ElseIf txtCodigoProprietario4.Value.Length > 0 AndAlso txtRNTRC4.Text.Length = 0 Then
            MsgBox(Me.Page, "RNTRC da Quarta Placa não foi informado.")
            Return False
        ElseIf radSim.Checked Then
            MsgBox(Me.Page, "Placa com restrição não pode ser utilizada.")
            Return False
        ElseIf Not String.IsNullOrEmpty(txtPlaca1.Text) AndAlso String.IsNullOrEmpty(txtPlaca2.Text) AndAlso (Not String.IsNullOrEmpty(txtPlaca3.Text) OrElse Not String.IsNullOrEmpty(txtPlaca4.Text)) Then
            MsgBox(Me.Page, "A Placa 2 deve ser preenchida antes da placa 3 ou placa 4 .")
            Return False
        ElseIf Not String.IsNullOrEmpty(txtPlaca1.Text) AndAlso Not String.IsNullOrEmpty(txtPlaca2.Text) AndAlso String.IsNullOrEmpty(txtPlaca3.Text) AndAlso Not String.IsNullOrEmpty(txtPlaca4.Text) Then
            MsgBox(Me.Page, "A placa 2 e placa 3 devem ser preenchidas antes da placa 4.")
            Return False
        ElseIf ddlViaDeTransporte.SelectedValue.Equals("1") AndAlso Not txtPlaca1.Text = "BONIFICA" Then 'Rodoviário
            If (Not String.IsNullOrEmpty(txtPlaca1.Text) AndAlso (Not ValidarPlaca(txtPlaca1.Text) OrElse txtPlaca1.Text.Trim().Replace("-", "").Length < 6)) Then
                MsgBox(Me.Page, "Placa 1 Inválida!")
                Return False
            ElseIf (Not String.IsNullOrEmpty(txtPlaca2.Text) AndAlso (Not ValidarPlaca(txtPlaca2.Text) OrElse txtPlaca2.Text.Trim().Replace("-", "").Length < 6)) Then
                MsgBox(Me.Page, "Placa 2 Inválida!")
                Return False
            ElseIf (Not String.IsNullOrEmpty(txtPlaca3.Text) AndAlso (Not ValidarPlaca(txtPlaca3.Text) OrElse txtPlaca3.Text.Trim().Replace("-", "").Length < 6)) Then
                MsgBox(Me.Page, "Placa 3 Inválida!")
                Return False
            ElseIf (Not String.IsNullOrEmpty(txtPlaca4.Text) AndAlso (Not ValidarPlaca(txtPlaca4.Text) OrElse txtPlaca4.Text.Trim().Replace("-", "").Length < 6)) Then
                MsgBox(Me.Page, "Placa 4 Inválida!")
                Return False
            End If
        End If
        Return True
    End Function

    Protected Sub btnProprietario1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnProprietario1.Click
        If String.IsNullOrWhiteSpace(txtPlaca1.Text) Then
            MsgBox(Me.Page, "A primeira placa não foi informada!")
        Else
            Dim ucConsultaClientes As ucConsultaClientes = DirectCast(Me.NamingContainer.FindControl("ucConsultaClientes"), ucConsultaClientes)
            If ucConsultaClientes IsNot Nothing Then
                ucConsultaClientes.Limpar()
                ucConsultaClientes.SetarTipoCliente("7,13")
                ucConsultaClientes.SetarTituloDIV("Consulta de Proprietário")
                ucConsultaClientes.MainUserControl = Me
                Popup.ConsultaDeClientes(Me.Page, "objProprietarioPLXP1" & HID.Value)
            End If
        End If
    End Sub

    Protected Sub btnProprietario2_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If String.IsNullOrWhiteSpace(txtPlaca2.Text) Then
            MsgBox(Me.Page, "A segunda placa não foi informada!")
        Else
            Dim ucConsultaClientes As ucConsultaClientes = DirectCast(Me.NamingContainer.FindControl("ucConsultaClientes"), ucConsultaClientes)
            If ucConsultaClientes IsNot Nothing Then
                ucConsultaClientes.Limpar()
                ucConsultaClientes.SetarTipoCliente("7,13")
                ucConsultaClientes.SetarTituloDIV("Consulta de Proprietário")
                ucConsultaClientes.MainUserControl = Me
            End If
            Popup.ConsultaDeClientes(Me.Page, "objProprietarioPLXP2" & HID.Value.ToString)
        End If
    End Sub

    Protected Sub btnProprietario3_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If String.IsNullOrWhiteSpace(txtPlaca3.Text) Then
            MsgBox(Me.Page, "A terceira placa não foi informada!")
        Else
            Dim ucConsultaClientes As ucConsultaClientes = DirectCast(Me.NamingContainer.FindControl("ucConsultaClientes"), ucConsultaClientes)
            If ucConsultaClientes IsNot Nothing Then
                ucConsultaClientes.Limpar()
                ucConsultaClientes.SetarTipoCliente("7,13")
                ucConsultaClientes.SetarTituloDIV("Consulta de Proprietário")
                ucConsultaClientes.MainUserControl = Me
            End If
            Popup.ConsultaDeClientes(Me.Page, "objProprietarioPLXP3" & HID.Value.ToString)
        End If
    End Sub

    Protected Sub btnProprietario4_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If String.IsNullOrWhiteSpace(txtPlaca4.Text) Then
            MsgBox(Me.Page, "A quarta placa não foi informada!")
        Else
            Dim ucConsultaClientes As ucConsultaClientes = DirectCast(Me.NamingContainer.FindControl("ucConsultaClientes"), ucConsultaClientes)
            If ucConsultaClientes IsNot Nothing Then
                ucConsultaClientes.Limpar()
                ucConsultaClientes.SetarTipoCliente("7,13")
                ucConsultaClientes.SetarTituloDIV("Consulta de Proprietário")
                ucConsultaClientes.MainUserControl = Me
            End If
            Popup.ConsultaDeClientes(Me.Page, "objProprietarioPLXP4" & HID.Value.ToString)
        End If
    End Sub

    Protected Sub imgConsultaRNTRC2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        If txtProprietario2.Text.Length = 0 Then
            MsgBox(Me.Page, "Proprietário da segunda Placa não foi informado.")
        ElseIf txtRNTRC2.Text.Length = 0 Then
            MsgBox(Me.Page, "RNTRC da segunda Placa não foi informado.")
        ElseIf txtPlaca2.Text.Length = 0 Then
            MsgBox(Me.Page, "Segunda Placa não foi informada.")
        Else
            Dim objPlaca As New [Lib].Negocio.Placa()
            objPlaca.Placa01 = Trim(txtPlaca2.Text).Replace("-", "")

            Dim ds As New DataSet
            If ds Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
                MsgBox(Me.Page, "Erro inesperado consultando Frota, entre em contato com o Suporte")
            Else
                If ds.Tables(0).Rows(0).Item("erro") = 0 Then
                    MsgBox(Me.Page, ds.Tables(0).Rows(0).Item("errodescricao"))
                Else
                    MsgBox(Me.Page, "Erro: Código " & ds.Tables(0).Rows(0).Item("erro") & " - " & ds.Tables(0).Rows(0).Item("errodescricao"))
                End If
            End If
        End If
    End Sub

    Protected Sub imgConsultaRNTRC3_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        If txtProprietario3.Text.Length = 0 Then
            MsgBox(Me.Page, "Proprietário da terceira Placa não foi informado.")
        ElseIf txtRNTRC3.Text.Length = 0 Then
            MsgBox(Me.Page, "RNTRC da terceira Placa não foi informado.")
        ElseIf txtPlaca3.Text.Length = 0 Then
            MsgBox(Me.Page, "Terceira Placa não foi informada.")
        Else
            Dim objPlaca As New [Lib].Negocio.Placa()
            objPlaca.Placa01 = Trim(txtPlaca3.Text).Replace("-", "")

            Dim ds As New DataSet
            If ds Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
                MsgBox(Me.Page, "Erro inesperado consultando Frota, entre em contato com o Suporte")
            Else
                If ds.Tables(0).Rows(0).Item("erro") = 0 Then
                    MsgBox(Me.Page, ds.Tables(0).Rows(0).Item("errodescricao"))
                Else
                    MsgBox(Me.Page, "Erro: Código " & ds.Tables(0).Rows(0).Item("erro") & " - " & ds.Tables(0).Rows(0).Item("errodescricao"))
                End If
            End If
        End If
    End Sub

    Protected Sub imgConsultaRNTRC4_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        If txtProprietario4.Text.Length = 0 Then
            MsgBox(Me.Page, "Proprietário da quarta Placa não foi informado.")
        ElseIf txtRNTRC4.Text.Length = 0 Then
            MsgBox(Me.Page, "RNTRC da quarta Placa não foi informado.")
        ElseIf txtPlaca4.Text.Length = 0 Then
            MsgBox(Me.Page, "Quarta Placa não foi informada.")
        Else
            Dim objPlaca As New [Lib].Negocio.Placa()
            objPlaca.Placa01 = Trim(txtPlaca4.Text).Replace("-", "")

            Dim ds As New DataSet
            If ds Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
                MsgBox(Me.Page, "Erro inesperado consultando Frota, entre em contato com o Suporte")
            Else
                If ds.Tables(0).Rows(0).Item("erro") = 0 Then
                    MsgBox(Me.Page, ds.Tables(0).Rows(0).Item("errodescricao"))
                Else
                    MsgBox(Me.Page, "Erro: Código " & ds.Tables(0).Rows(0).Item("erro") & " - " & ds.Tables(0).Rows(0).Item("errodescricao"))
                End If
            End If
        End If
    End Sub

    Protected Sub btnRepPro2_Click(ByVal sender As Object, ByVal e As System.EventArgs)
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

    Protected Sub btnRepPro3_Click(ByVal sender As Object, ByVal e As System.EventArgs)
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

    Protected Sub btnRepPro4_Click(ByVal sender As Object, ByVal e As System.EventArgs)
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

    Protected Sub imgLimparPlaca2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imgLimparPlaca2.Click
        Try
            LimparPlaca2()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgLimparPlaca3_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imgLimparPlaca3.Click
        Try
            LimparPlaca3()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try

    End Sub

    Protected Sub imgLimparPlaca4_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imgLimparPlaca4.Click
        Try
            LimparPlaca4()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
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

    Private Sub Bloquear(ByVal enabled As Boolean)
        txtPlaca1.Enabled = Not enabled
        txtRNTRC1.Enabled = enabled
        btnProprietario1.Enabled = enabled

        txtPlaca2.Enabled = enabled
        txtRNTRC2.Enabled = enabled
        btnProprietario2.Enabled = enabled

        txtPlaca3.Enabled = enabled
        txtRNTRC3.Enabled = enabled
        btnProprietario3.Enabled = enabled

        txtPlaca4.Enabled = enabled
        txtRNTRC4.Enabled = enabled
        btnProprietario4.Enabled = enabled

        imgLimparPlaca2.Enabled = enabled
        btnProprietario2.Enabled = enabled
        btnRepPro2.Enabled = enabled

        imgLimparPlaca3.Enabled = enabled
        btnProprietario3.Enabled = enabled
        btnRepPro3.Enabled = enabled

        imgLimparPlaca4.Enabled = enabled
        btnProprietario4.Enabled = enabled
        btnRepPro4.Enabled = enabled

        btnMotorista.Enabled = enabled
        imgLimparMotorista.Enabled = enabled

        ddlTipoDeVeiculo.Enabled = enabled
        ddlViaDeTransporte.Enabled = enabled
        btnMotorista.Enabled = enabled
        imgLimparMotorista.Enabled = enabled
        txtHabilitacao.Enabled = enabled
    End Sub

    Protected Sub lnkConfirmar_Click(sender As Object, e As EventArgs) Handles lnkConfirmar.Click
        Try
            Confirmar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkFechar_Click(sender As Object, e As EventArgs) Handles lnkFechar.Click
        If MainUserControl IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(MainUserControl.ClientID) Then
            Popup.CloseDialog(Me.Page, "divConsultaPlacas")
        Else
            objPlaca = New [Lib].Negocio.Placa()
            CType(Me.Page, IBasePage).Carregar(objPlaca)
            Popup.CloseDialog(Me.Page, "divConsultaPlacas")
        End If
    End Sub

    Protected Sub btnMotorista_Click(sender As Object, e As EventArgs) Handles btnMotorista.Click
        Dim ucConsultaClientes As ucConsultaClientes = DirectCast(Me.NamingContainer.FindControl("ucConsultaClientes"), ucConsultaClientes)
        If ucConsultaClientes IsNot Nothing Then
            ucConsultaClientes.Limpar()
            ucConsultaClientes.SetarTipoCliente("13")
            ucConsultaClientes.SetarTituloDIV("Consulta de Motorista")
            ucConsultaClientes.MainUserControl = Me
        End If
        Popup.ConsultaDeClientes(Me.Page, "objMotoristaPXE" & HID.Value.ToString)
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Consultar()
    End Sub

    Private Sub Consultar()
        If String.IsNullOrWhiteSpace(txtPlaca1.Text) Then
            MsgBox(Me.Page, "Informe o número da placa.")
            txtPlaca1.Focus()
            Exit Sub
        End If

        Dim Placa As String = String.Empty

        If Left(Session("ssEmpresa"), 8) = "05272759" Then
            Placa = txtPlaca1.Text
        Else
            Placa = Funcoes.EliminarCaracteresEspeciais(txtPlaca1.Text)
        End If

        Placa = Placa.Replace(".", "").Replace(",", "").Replace(" ", "")
        txtPlaca1.Text = Placa
        txtPlaca1.Enabled = False
        objPlaca = New [Lib].Negocio.Placa(txtPlaca1.Text)

        If objPlaca.Selecionar = False Then
            If Not objPlaca.Erro Is Nothing Then
                MsgBox(Me.Page, objPlaca.Erro.ToString, eTitulo.Info)
            Else
                MsgBox(Me.Page, "Placa não encontrada.")
            End If
        Else
            If objPlaca.CodigoProprietario01.Length = 0 AndAlso Not CType(Session("objNotaFiscal" & HID.Value.ToString), [Lib].Negocio.NotaFiscal) Is Nothing Then
                objPlaca.CodigoProprietario01 = CType(Session("objNotaFiscal" & HID.Value.ToString), [Lib].Negocio.NotaFiscal).CodigoTransportador
                objPlaca.EndProprietario01 = CType(Session("objNotaFiscal" & HID.Value.ToString), [Lib].Negocio.NotaFiscal).EnderecoTransportador
                objPlaca.Proprietario01.Nome = Trim(CType(Session("objNotaFiscal" & HID.Value.ToString), [Lib].Negocio.NotaFiscal).Transportador.Nome)
                If CType(Session("objNotaFiscal" & HID.Value.ToString), [Lib].Negocio.NotaFiscal).Transportador.RNTRCTransportador.Length > 0 AndAlso objPlaca.RNTRCPlaca01.Length = 0 Then
                    objPlaca.RNTRCPlaca01 = CType(Session("objNotaFiscal" & HID.Value.ToString), [Lib].Negocio.NotaFiscal).Transportador.RNTRCTransportador
                    txtRNTRC1.Text = objPlaca.RNTRCPlaca01
                End If
            End If

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

            ddlTipoDeVeiculo.SelectedValue = objPlaca.TipoDeVeiculoDetalhes.Codigo
            ddlViaDeTransporte.SelectedValue = objPlaca.ViaDeTransporteDetalhes.Codigo

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

            If Not String.IsNullOrWhiteSpace(objPlaca.Observacao) Then
                txtObservacao.Enabled = True
                txtObservacao.Text = objPlaca.Observacao
            End If

            If objPlaca.Restricao = "S" Then
                radSim.Checked = True
            Else
                radNao.Checked = True
            End If
        End If

        lnkConfirmar.Parent.Visible = True
        Bloquear(True)
        SalvarVariavelPlaca()
    End Sub
End Class