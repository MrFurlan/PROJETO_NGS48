Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class BalancasEletronicas
    Inherits BasePage

    Dim Mensagem As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("BalancasEletronicas", "ACESSAR") Then
                Limpar()
                ConsultarBalancas()
                CarregarUsuarios()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CarregarUsuarios()
        ddl.Carregar(ddlUsuario, CarregarDDL.Tabela.Usuarios, "", True)
    End Sub

    Private Sub Limpar()
        Session.Remove("objBalanca" & HID.Value)
        Dim objBalanca As New BalancaEletronica()
        HID.Value = Guid.NewGuid().ToString
        Session("objBalanca" & HID.Value) = objBalanca
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
        txtNomeBalanca.Enabled = True
        txtIPBalanca.Enabled = True
        ddlModelo.Enabled = True
        radESim.Checked = False
        radENao.Checked = False
        txtNomeBalanca.Text = ""
        txtIPBalanca.Text = ""
        txtReadBuffer.Text = ""
        txtReceiveBytes.Text = ""
        txtBaudRate.Text = ""
        txtDataBits.Text = ""
        txtStopBits.Text = ""
        ConsultarBalancas()
        lnkNovoU.Parent.Visible = True
        lnkExcluirU.Parent.Visible = False
        lnkLimparU.Parent.Visible = True
        ddlUsuario.Enabled = False
        txtIpUsuario.Enabled = False
        ddlUsuario.SelectedIndex = 0
        txtIpUsuario.Text = ""
        gridBalancaXUsuarios.DataSource = Nothing
        gridBalancaXUsuarios.DataBind()
    End Sub

    Private Sub LimparUsuarios()
        ddlUsuario.Enabled = False
        txtIpUsuario.Enabled = False
        lnkNovoU.Parent.Visible = True
        lnkExcluirU.Parent.Visible = False
        lnkLimparU.Parent.Visible = True
        ddlUsuario.SelectedIndex = 0
        txtIpUsuario.Text = ""
    End Sub

    Private Sub ConsultarBalancas()
        If Funcoes.VerificaPermissao("BalancasEletronicas", "LEITURA") Then
            Dim lst As New ListBalancaEletronica()
            gridBalancaEletronica.DataSource = lst.ToArray()
            gridBalancaEletronica.DataBind()
        End If
    End Sub

    Private Sub CarregarBalancasXUsuarios(ByVal Balanca As String)
        Dim lst As New ListBalancasXUsuarios(Balanca)
        gridBalancaXUsuarios.DataSource = lst.ToArray()
        gridBalancaXUsuarios.DataBind()
    End Sub

    Function ValidarCampos() As Boolean
        If txtNomeBalanca.Text.Length = 0 Then
            Mensagem = "Nome da Balança não foi informado"
            Return False
        ElseIf txtIPBalanca.Text.Length = 0 Then
            Mensagem = "IP da Balança não foi informado"
            Return False
        ElseIf radESim.Checked = False And radENao.Checked = False Then
            Mensagem = "Informe se Balança é Eltrónica"
            Return False
        ElseIf txtReadBuffer.Text.Length = 0 Then
            Mensagem = "Read Buffer não foi informado"
            Return False
        ElseIf txtReceiveBytes.Text.Length = 0 Then
            Mensagem = "Receive Bytes não foi informado"
            Return False
        ElseIf txtBaudRate.Text.Length = 0 Then
            Mensagem = "Baud Rate não foi informado"
            Return False
        ElseIf txtDataBits.Text.Length = 0 Then
            Mensagem = "DataBits não foi informado"
            Return False
        ElseIf txtStopBits.Text.Length = 0 Then
            Mensagem = "StopBits não foi informado"
            Return False
        Else
            Return True
        End If
    End Function

    Private Sub IniciarProcesso(ByVal Tipo As String)
        If Tipo = "I" And Funcoes.VerificaPermissao("BalancasEletronicas", "GRAVAR") Or _
           Tipo = "U" And Funcoes.VerificaPermissao("BalancasEletronicas", "ALTERAR") Or _
           Tipo = "D" And Funcoes.VerificaPermissao("BalancasEletronicas", "EXCLUIR") Then

            If ValidarCampos() Then
                CType(Session("objBalanca" & HID.Value), BalancaEletronica).IUD = Tipo
                CType(Session("objBalanca" & HID.Value), BalancaEletronica).BalancaServidor = txtNomeBalanca.Text
                CType(Session("objBalanca" & HID.Value), BalancaEletronica).BalancaIp = txtIPBalanca.Text
                CType(Session("objBalanca" & HID.Value), BalancaEletronica).BalancaTipo = ddlModelo.SelectedValue
                If radESim.Checked Then
                    CType(Session("objBalanca" & HID.Value), BalancaEletronica).Eletronica = "S"
                ElseIf radENao.Checked Then
                    CType(Session("objBalanca" & HID.Value), BalancaEletronica).Eletronica = "N"
                End If
                CType(Session("objBalanca" & HID.Value), BalancaEletronica).ReadBuffer = txtReadBuffer.Text
                CType(Session("objBalanca" & HID.Value), BalancaEletronica).ReceivedBytes = txtReceiveBytes.Text
                CType(Session("objBalanca" & HID.Value), BalancaEletronica).PortName = ddlPorta.SelectedValue
                CType(Session("objBalanca" & HID.Value), BalancaEletronica).BaudRate = txtBaudRate.Text
                CType(Session("objBalanca" & HID.Value), BalancaEletronica).Parity = ddlParidade.SelectedValue
                CType(Session("objBalanca" & HID.Value), BalancaEletronica).DataBits = txtDataBits.Text
                CType(Session("objBalanca" & HID.Value), BalancaEletronica).StopBits = txtStopBits.Text

                If CType(Session("objBalanca" & HID.Value), BalancaEletronica).Salvar Then
                    Limpar()
                    ConsultarBalancas()
                Else
                    Select Case Tipo
                        Case "I"
                            MsgBox(Me.Page, "Erro ao Incluir Registro: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
                        Case "U"
                            MsgBox(Me.Page, "Erro ao Alterar Registro: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
                        Case "D"
                            MsgBox(Me.Page, "Erro ao Excluir Registro: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
                    End Select
                End If
            Else
                MsgBox(Me.Page, Mensagem)
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para alterar Registro")
        End If
    End Sub

    Private Sub IniciarProcessoUsuario(ByVal Tipo As String)
        If txtIpUsuario.Text.Length = 0 Then
            MsgBox(Me.Page, "IP do Usuário não foi informado")
        ElseIf ddlUsuario.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Nome do Usuário não foi selecionado")
        Else
            Dim objBalancaXUsuario As New BalancasXUsuarios
            objBalancaXUsuario.IUD = Tipo
            objBalancaXUsuario.CodigoBalancaServidor = CType(Session("objBalanca" & HID.Value), BalancaEletronica).BalancaServidor
            objBalancaXUsuario.BalancaUsuarioIp = txtIpUsuario.Text
            objBalancaXUsuario.NomeUsuario = ddlUsuario.SelectedValue

            If objBalancaXUsuario.Salvar Then
                CarregarBalancasXUsuarios(CType(Session("objBalanca" & HID.Value), BalancaEletronica).BalancaServidor)
                LimparUsuarios()
            Else
                Select Case Tipo
                    Case "I"
                        MsgBox(Me.Page, "Erro ao Incluir Registro: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
                    Case "D"
                        MsgBox(Me.Page, "Erro ao Excluir Registro: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
                End Select
            End If
        End If
    End Sub

    Protected Sub gridBalancaEletronica_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim objBalanca As New BalancaEletronica(gridBalancaEletronica.SelectedRow.Cells.Item(1).Text())
            Session("objBalanca" & HID.Value) = objBalanca

            txtNomeBalanca.Text = objBalanca.BalancaServidor
            txtIPBalanca.Text = objBalanca.BalancaIp
            ddlModelo.SelectedValue = objBalanca.BalancaTipo
            If objBalanca.Eletronica = "S" Then
                radESim.Checked = True
                radENao.Checked = False
            ElseIf objBalanca.Eletronica = "N" Then
                radENao.Checked = True
                radESim.Checked = False
            End If
            txtReadBuffer.Text = objBalanca.ReadBuffer
            txtReceiveBytes.Text = objBalanca.ReceivedBytes
            ddlPorta.SelectedValue = objBalanca.PortName
            txtBaudRate.Text = objBalanca.BaudRate
            ddlParidade.SelectedValue = objBalanca.Parity
            txtDataBits.Text = objBalanca.DataBits
            txtStopBits.Text = objBalanca.StopBits
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
            txtNomeBalanca.Enabled = False
            txtIPBalanca.Enabled = False
            ddlModelo.Enabled = False
            LimparUsuarios()
            ddlUsuario.Enabled = True
            txtIpUsuario.Enabled = True
            CarregarBalancasXUsuarios(objBalanca.BalancaServidor)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnLimpar_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Limpar()
            ConsultarBalancas()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub gridBalancaXUsuarios_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            lnkNovoU.Parent.Visible = False
            lnkExcluirU.Parent.Visible = True
            ddlUsuario.Enabled = False
            txtIpUsuario.Enabled = False
            ddlUsuario.SelectedValue = gridBalancaXUsuarios.SelectedRow.Cells(3).Text()
            txtIpUsuario.Text = gridBalancaXUsuarios.SelectedRow.Cells(2).Text()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            IniciarProcesso("I")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click
        Try
            IniciarProcesso("U")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        Try
            IniciarProcesso("D")
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
            Funcoes.Ajuda(Me.Page, "BalancasEletronicas")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub lnkNovoU_Click(sender As Object, e As EventArgs) Handles lnkNovoU.Click
        Try
            IniciarProcessoUsuario("I")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluirU_Click(sender As Object, e As EventArgs) Handles lnkExcluirU.Click
        Try
            IniciarProcessoUsuario("D")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimparU_Click(sender As Object, e As EventArgs) Handles lnkLimparU.Click
        Try
            LimparUsuarios()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjudaU_Click(sender As Object, e As EventArgs) Handles lnkAjudaU.Click

    End Sub

End Class