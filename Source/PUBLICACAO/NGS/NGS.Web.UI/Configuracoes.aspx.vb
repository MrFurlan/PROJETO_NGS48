Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Configuracoes
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("Configuracoes", "ACESSAR") Then
                CarregarConfiguracao()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página!", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        If Funcoes.VerificaPermissao("Configuracoes", "GRAVAR") Then
            Dim objConfiguracoes As New Configuracao()
            If Not String.IsNullOrWhiteSpace(txtCodigo.Text) AndAlso CInt(txtCodigo.Text) > 0 Then objConfiguracoes.Codigo = txtCodigo.Text.Trim()
            If Not String.IsNullOrWhiteSpace(txtCodigo.Text) AndAlso CInt(txtCodigo.Text) > 0 Then objConfiguracoes.IUD = "U" Else objConfiguracoes.IUD = "I"
            objConfiguracoes.Email = txtEmail.Text.Trim()
            objConfiguracoes.Host = txtHost.Text.Trim()
            objConfiguracoes.Senha = txtSenha.Text.Trim()
            objConfiguracoes.Usuario = txtUsuario.Text.Trim()
            objConfiguracoes.Porta = txtPorta.Text.Trim()
            objConfiguracoes.Cheque = txtCheque.Text.Trim()
            objConfiguracoes.Balanca = txtBalanca.Text.Trim()
            objConfiguracoes.FilesServer = txtFilesServer.Text.Trim()
            objConfiguracoes.Credenciail = chkCredenciais.Checked
            objConfiguracoes.Ssl = chkSsl.Checked
            If objConfiguracoes.Salvar() Then
                MsgBox(Me.Page, "Configuração salva com Sucesso.", eTitulo.Sucess)
                CarregarConfiguracao()
            Else
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para salvar registro!")
        End If
    End Sub

    Protected Sub lnkSair_Click(sender As Object, e As EventArgs) Handles lnkSair.Click
        Response.Redirect("~/Gestao.aspx")
    End Sub

    Private Sub CarregarConfiguracao()
        Dim obj As New Configuracao()
        txtCodigo.Text = obj.Codigo
        txtEmail.Text = obj.Email
        txtHost.Text = obj.Host
        txtSenha.Text = obj.Senha
        txtUsuario.Text = obj.Usuario
        txtPorta.Text = obj.Porta
        txtCheque.Text = obj.Cheque
        txtBalanca.Text = obj.Balanca
        txtFilesServer.Text = obj.FilesServer
        chkCredenciais.Checked = obj.Credenciail
        chkSsl.Checked = obj.Ssl
        CarregarNaoVinculados()
        CarregarVinculados()
    End Sub

    Private Function getEtapa() As eEtapa
        If rdoCheque.Checked Then
            Return eEtapa.Cheque
        ElseIf rdoNFePendencias.Checked Then
            Return eEtapa.NFePendencias
        ElseIf rdoEstoqueMinimo.Checked Then
            Return eEtapa.EstoqueMinimo
        ElseIf rdDFe.Checked Then
            Return eEtapa.DFeSefaz
        ElseIf rdContabilizarNotas.Checked Then
            Return eEtapa.ContabilizarNotas
        End If
    End Function

    Private Sub CarregarNaoVinculados()
        Dim sql As String = "SELECT u.Usuario_Id, u.NomeCompleto, u.Email " & vbCrLf & _
                            "FROM Usuarios u " & vbCrLf & _
                            "WHERE 1=1 " & vbCrLf & _
                            "AND u.Usuario_Id NOT IN (SELECT Usuario_Id FROM ConfiguracaoXUsuario cxu WHERE cxu.Etapa_Id = " & CInt(getEtapa()) & ") " & vbCrLf & _
                            "AND u.Email IS NOT NULL " & vbCrLf & _
                            "AND u.Email <> '' " & vbCrLf & _
                            "AND u.Email <> '&nbsp;' "
        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "NaoVinculados")
        grdNaoVinculados.DataSource = ds
        grdNaoVinculados.DataBind()
    End Sub

    Private Sub CarregarVinculados()
        Dim sql As String = "SELECT u.Usuario_Id, u.NomeCompleto, u.Email " & vbCrLf & _
                            "FROM Usuarios u " & vbCrLf & _
                            "WHERE 1=1 " & vbCrLf & _
                            "AND u.Usuario_Id IN (SELECT Usuario_Id FROM ConfiguracaoXUsuario cxu WHERE cxu.Etapa_Id = " & CInt(getEtapa()) & ") " & vbCrLf & _
                            "AND u.Email IS NOT NULL " & vbCrLf & _
                            "AND u.Email <> '' " & vbCrLf & _
                            "AND u.Email <> '&nbsp;' "
        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Vinculados")
        grdVinculados.DataSource = ds
        grdVinculados.DataBind()
    End Sub

    Protected Sub lnkDireita_Click(sender As Object, e As EventArgs) Handles lnkDireita.Click
        If CInt(txtCodigo.Text) > 0 Then
            Dim lstSelecionados As List(Of String) = grdNaoVinculados.GetSelectedValues("chkSelecionado")
            If lstSelecionados IsNot Nothing AndAlso lstSelecionados.Count > 0 Then
                For Each strUsuario As String In lstSelecionados
                    Dim sql As String = "INSERT INTO ConfiguracaoXUsuario (Configuracao_Id, Etapa_Id, Usuario_Id) " & vbCrLf & _
                                        "VALUES ('" & txtCodigo.Text.Trim() & "', '" & CInt(getEtapa()) & "', '" & strUsuario & "');"
                    If Not Banco.GravaBanco(sql) Then
                        MsgBox(Me.Page, Session("ssMessage"))
                    End If
                Next
                CarregarNaoVinculados()
                CarregarVinculados()
            End If
        Else
            MsgBox(Me.Page, "É necessário salvar a configuração antes de vincular o(s) e-mail(s)!")
        End If
    End Sub

    Protected Sub lnkEsquerda_Click(sender As Object, e As EventArgs) Handles lnkEsquerda.Click
        If CInt(txtCodigo.Text) > 0 Then
            Dim lstSelecionados As List(Of String) = grdVinculados.GetSelectedValues("chkSelecionado")
            If lstSelecionados IsNot Nothing AndAlso lstSelecionados.Count > 0 Then
                For Each strUsuario As String In lstSelecionados
                    Dim sql As String = "DELETE FROM ConfiguracaoXUsuario " & vbCrLf & _
                                        "WHERE Configuracao_Id = '" & txtCodigo.Text.Trim() & "' AND Etapa_Id = '" & CInt(getEtapa()) & "' AND Usuario_Id = '" & strUsuario & "';"
                    If Not Banco.GravaBanco(sql) Then
                        MsgBox(Me.Page, Session("ssMessage"))
                    End If
                Next
                CarregarNaoVinculados()
                CarregarVinculados()
            End If
        Else
            MsgBox(Me.Page, "É necessário salvar a configuração antes de desvincular o(s) e-mail(s)!")
        End If
    End Sub

    Protected Sub rdoCheque_CheckedChanged(sender As Object, e As EventArgs) Handles rdoCheque.CheckedChanged
        CarregarNaoVinculados()
        CarregarVinculados()
    End Sub

    Protected Sub rdoNFePendencias_CheckedChanged(sender As Object, e As EventArgs) Handles rdoNFePendencias.CheckedChanged
        CarregarNaoVinculados()
        CarregarVinculados()
    End Sub

    Protected Sub rdoEstoqueMinimo_CheckedChanged(sender As Object, e As EventArgs) Handles rdoEstoqueMinimo.CheckedChanged
        CarregarNaoVinculados()
        CarregarVinculados()
    End Sub

    Protected Sub rdDFe_CheckedChanged(sender As Object, e As EventArgs) Handles rdDFe.CheckedChanged
        CarregarNaoVinculados()
        CarregarVinculados()
    End Sub

    Protected Sub rdContabilizarNotas_CheckedChanged(sender As Object, e As EventArgs) Handles rdContabilizarNotas.CheckedChanged
        CarregarNaoVinculados()
        CarregarVinculados()
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "Configuracoes")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class