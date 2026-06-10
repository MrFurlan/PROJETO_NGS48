Imports System.IO
Imports System.Net
Imports System.Reflection
Imports System.Web.Configuration
Imports System.Xml
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Login
    Inherits BasePage

    Structure Noticias
        Public Property Titulo As String
        Public Property Data As DateTime
        Public Property Descricao As String
        Public Property Link As String
    End Structure

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            pageLoad()
        End If
    End Sub

    Protected Sub btnEntrar_Click(ByVal sender As Object, ByVal e As EventArgs)
        Try
            Entrar()
        Catch ex As Exception
            cmbBancoServidor.Visible = False
            lblErro.Text = ex.Message
            lblErro.Visible = True
        End Try
    End Sub

    Protected Sub btnRecuperarSenha_Click(ByVal sender As Object, ByVal e As EventArgs)
        loginMenu.Visible = False
        recuperarSenhaMenu.Visible = True

        Dim lstBancos = BancoDados.ListarBancoPorUsuarios(txtUsuario.Text)

        For Each obj In lstBancos
            cmbBancoPorUsuario.Items.Add(New ListItem(String.Format("{0} | {1}", obj.Banco_Id.Trim().PadRight(20 - obj.Banco_Id.Length, "."), obj.HostServidor),
                                                    String.Format("{0}|{1}", obj.Banco_Id.Trim(), obj.HostServidor)))
        Next
    End Sub

    Protected Sub btnEnviarEmail_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim sUsuario As String = txtUsuarioRecuperacao.Text
        Dim objRecuperarSenha As RecuperarSenha = New RecuperarSenha(txtUsuario.Text)

    End Sub

    Protected Sub btnVoltar_Click(ByVal sender As Object, ByVal e As EventArgs)
        Limpar()
    End Sub

    Private Sub pageLoad()
        Dim objAssembly As Assembly = Assembly.GetExecutingAssembly()
        Dim objName As AssemblyName = objAssembly.GetName()
        Dim objVersao As Version = objName.Version

        lblVersao.Text = objVersao.ToString()

        Limpar()
        If txtUsuario IsNot Nothing Then
            txtUsuario.Focus()
            'txtUsuario.Text = "TESTE"
            'txtSenha.Text = "123"
            'txtUsuario.Enabled = False
            'txtSenha.Enabled = False
        End If

        If Not Session("Acontecimento") Is Nothing Then
            divMensagem.Visible = True
            Select Case CType(Session("Acontecimento"), AcontecimentosGerais)
                Case AcontecimentosGerais.SenhaModificada
                    lblMsg.Text = "A senha foi modificada com sucesso!"
            End Select
        End If
    End Sub

    Private Sub Entrar()
        lblErro.Text = String.Empty
        lblErro.Visible = False
        Dim lstBancos As List(Of BancoDados) = BancoDados.ListarBancoPorUsuarios(txtUsuario.Text)


        If Request.Url.Host.ToUpper.Contains("LOCALHOST") AndAlso Not cmbBancoServidor.Visible AndAlso String.IsNullOrWhiteSpace(txtSenha.Text) Then
            cmbBancoServidor.Visible = True
            lstBancos = BancoDados.Listar()

            For Each obj In lstBancos
                cmbBancoServidor.Items.Add(New ListItem(String.Format("{0} | {1}", obj.Banco_Id.Trim().PadRight(20 - obj.Banco_Id.Length, "."), obj.HostServidor),
                                                        String.Format("{0}|{1}", obj.Banco_Id.Trim(), obj.HostServidor)))
            Next
        ElseIf lstBancos.Count > 1 And Not cmbBancoServidor.Visible AndAlso String.IsNullOrWhiteSpace(txtSenha.Text) Then
            cmbBancoServidor.Visible = True

            For Each obj In lstBancos
                cmbBancoServidor.Items.Add(New ListItem(String.Format("{0} | {1}", obj.Banco_Id.Trim().PadRight(20 - obj.Banco_Id.Length, "."), obj.HostServidor),
                                                        String.Format("{0}|{1}", obj.Banco_Id.Trim(), obj.HostServidor)))
            Next
        ElseIf String.IsNullOrWhiteSpace(txtUsuario.Text) OrElse String.IsNullOrWhiteSpace(txtSenha.Text) Then
            lblErro.Text = "Informe o usuário e a senha!"
            lblErro.Visible = True
            txtSenha.Focus()
        Else
            Try
                If cmbBancoServidor.Visible = True Then
                    Dim strValor As String() = cmbBancoServidor.SelectedValue.Split("|")
                    UsuarioServidor.Conexao = UsuarioLocal.GetStringConexao(txtUsuario.Text, UsuarioServidor.NomeServidor, strValor(0), strValor(1))
                Else
                    UsuarioServidor.Conexao = UsuarioLocal.GetStringConexao(txtUsuario.Text, UsuarioServidor.NomeServidor)
                End If

                Dim strConexao As String() = UsuarioServidor.Conexao.ToString.Split(New Char() {";"}, StringSplitOptions.RemoveEmptyEntries)


                If strConexao.Count > 0 Then

                    If strConexao(0).Replace("Data Source=", String.Empty).Trim() = "SRVMARINGA\SQLEXPR" Then
                        UsuarioServidor.NomeServidor = "SRVMARINGA"
                    Else
                        UsuarioServidor.NomeServidor = strConexao(0).Replace("Data Source=", String.Empty).Trim()
                    End If

                    UsuarioServidor.EnderecoLocal = UsuarioServidor.Conexao.ToString()
                    UsuarioServidor.BancoDeDados = strConexao(1).Replace("Initial Catalog=", String.Empty).Trim()

                    '*************************************************************************************************
                    '*********   PARA AMBIENTE DE HOMOLOGAÇÂO DO FINANCEIRO NOVO   ***********************************
                    '*************************************************************************************************
                    'If Not UsuarioServidor.BancoDeDados.ToUpper = "FINANCEIRO" Then
                    '    MsgBox(Me.Page, "Ambiente de Homologação do Financeiro! verifique o usuário utilizado", False)
                    '    Exit Sub
                    'End If

                    Response.Cookies("conexao").Value = FuncoesStrings.CodificarPara64Bits(UsuarioServidor.Conexao)

                    If String.IsNullOrEmpty(Request.Cookies("conexao").Value) Then
                        lblErro.Text = "Usuário não cadastrado para nenhum banco de dados!"
                        lblErro.Visible = True
                        Exit Sub
                    Else
                        If Not UsuarioServidor.Validar(txtUsuario.Text, txtSenha.Text) Then
                            lblErro.Text = "Usuário ou senha inválidos!"
                            lblErro.Visible = True
                            txtSenha.Select()
                            txtSenha.Focus()
                            Exit Sub
                        ElseIf UsuarioServidor.TrocarSenha(txtUsuario.Text, txtSenha.Text) Then
                            UsuarioServidor.NmUsuario = txtUsuario.Text
                            Response.Redirect("~/SenhaNova.aspx")
                        ElseIf Not UsuarioServidor.UsuarioAtivo Then
                            lblErro.Text = "Usuário desabilitado para uso do Sistema!"
                            lblErro.Visible = True
                            txtUsuario.Select()
                            txtUsuario.Focus()
                        Else
                            Dim url As String = IIf(String.IsNullOrWhiteSpace(UsuarioServidor.MenuDeAcesso), "~/Index.aspx", String.Format("~/{0}.aspx", UsuarioServidor.MenuDeAcesso))
                            Response.Redirect(IIf(Request.QueryString("url") Is Nothing, url, Request.QueryString("url")))
                        End If
                    End If
                Else
                    lblErro.Text = "Este usuário não existe no banco de dados de Usuários!"
                    lblErro.Visible = True
                    txtUsuario.Select()
                    txtUsuario.Focus()
                End If

            Catch ex As Exception
                Throw New Exception(ex.Message)
            End Try
        End If
    End Sub

    Private Sub Limpar()
        FormsAuthentication.SignOut()
        HttpContext.Current.Session.Abandon()
        HttpContext.Current.Session.Clear()
        HttpContext.Current.Session.RemoveAll()

        txtUsuarioRecuperacao.Text = String.Empty
        txtSenha.Text = String.Empty
        txtUsuario.Text = String.Empty
        cmbBancoServidor.Visible = False
        recuperarSenhaMenu.Visible = False
        loginMenu.Visible = True
    End Sub

End Class