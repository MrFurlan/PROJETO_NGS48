Imports System.Web.Mvc
Imports System.Web.Security
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports NGS.Web.Api.Models

Public Class ContaController
    Inherits Controller

    Private Shared ReadOnly repository As IProdutoDashboardRepositorio = New ProdutoDashboardRepositorio()

    <AllowAnonymous>
    Public Function Login(returnUrl As String) As ActionResult

        Dim model As New LoginViewModel With {.ReturnUrl = returnUrl}

        Dim produto As New ProdutoDashboard
        model.BancosDeDados = repository.ConsultarBancoDeDados

        Return View(model)
    End Function

    <HttpPost>
    <AllowAnonymous>
    Public Function Login(model As LoginViewModel) As ActionResult
        Try
            If Not model.BancoId Is Nothing AndAlso model.BancoId.Length > 0 Then
                Dim strValor As String() = model.BancoId.Split("|")
                UsuarioServidor.Conexao = UsuarioLocal.GetStringConexao(model.UsuarioId, UsuarioServidor.NomeServidor, strValor(0), strValor(1))
            Else
                UsuarioServidor.Conexao = UsuarioLocal.GetStringConexao(model.UsuarioId, UsuarioServidor.NomeServidor)
            End If

            Dim strConexao As String() = UsuarioServidor.Conexao.ToString.Split(New Char() {";"}, StringSplitOptions.RemoveEmptyEntries)

            If strConexao.Count > 0 Then
                UsuarioServidor.NomeServidor = strConexao(0).Replace("Data Source=", String.Empty).Trim()
                UsuarioServidor.EnderecoLocal = UsuarioServidor.Conexao.ToString()
                UsuarioServidor.BancoDeDados = strConexao(1).Replace("Initial Catalog=", String.Empty).Trim()

                ' Verificação de conexão
                If String.IsNullOrWhiteSpace(UsuarioServidor.Conexao) Then
                    ModelState.AddModelError("", "Usuário não cadastrado para nenhum banco de dados!")
                    Return View(model)
                End If

                ' Define o cookie corretamente
                Dim cookie As New HttpCookie("conexao", FuncoesStrings.CodificarPara64Bits(UsuarioServidor.Conexao))
                Response.SetCookie(cookie)

                ' Validação do usuário
                If Not UsuarioServidor.Validar(model.UsuarioId, model.Password) Then
                    ModelState.AddModelError("", "Usuário ou senha inválidos!")
                    Return View(model)
                ElseIf Not UsuarioServidor.UsuarioAtivo Then
                    ModelState.AddModelError("", "Usuário desabilitado para uso do Sistema!")
                    Return View(model)
                Else
                    FormsAuthentication.SetAuthCookie(model.UsuarioId, False)
                    Return RedirectToAction("Index", "Dashboard")
                End If
            Else
                ModelState.AddModelError("", "Este usuário não existe no banco de dados de Usuários!")
                Return View(model)
            End If

        Catch ex As Exception
            ModelState.AddModelError("", ex.Message)
            Return View(model)
        End Try
    End Function

    Public Function Logout() As ActionResult

        FormsAuthentication.SignOut()
        Session.Abandon()

        ' (Opcional) Mensagem para exibir na tela de login
        TempData("MensagemLogout") = "Você saiu do sistema com sucesso."

        Return RedirectToAction("Login", "Conta")

    End Function

End Class
