<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Login.aspx.vb" Inherits="NGS.Web.UI.Login" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>NGS Soluções Ltda</title>
    <link id="favicon" rel="shortcut icon" href="Images/favicon.png" />
    <style type="text/css">
        .botaologin {
            background-color: #4D90FE;
            background-image: -moz-linear-gradient(center top, #4D90FE, #4787ED);
            border: 1px solid #3079ED;
            color: #FFFFFF;
            text-shadow: 0 1px rgba(0, 0, 0, 0.1);
            font-size: 12pt;
            height: 38px !important;
            margin: 0 1.5em 1.2em 0;
        }

        #box {
            height: 200px;
            left: 50%;
            margin: -100px 0 0 -200px;
            position: absolute;
            top: 50%;
            width: 400px;
        }

        input[type="text"]:focus, input[type="password"]:focus {
            background-color: #feffef;
        }
    </style>
    <script type="text/javascript">
        function click(event) {
            if (event.button == 2 || event.button == 3) {
                oncontextmenu = 'return false';
                //window.alert('Não disponível.');
            }
        }


        localStorage.clear();


        document.onmousedown = click
        document.oncontextmenu = new Function("return false;")

    </script>
</head>
<body>
    <form runat="server" autocomplete="off">
        <asp:ScriptManager ID="scrm" runat="server" />
        <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
        <div id="divMensagem" runat="server" class="mensagem" visible="false">
            <asp:Label ID="lblMsg" runat="server" Width="200px" Font-Bold="true" />
        </div>
        <div id="wrap" style="background: none repeat scroll 0 0 #FFFFFF !important;">
            <div id="topo">
                <div id="topomeio">
                    <asp:Image ID="imgLogo" runat="server" SkinID="imageLogo" AlternateText="Logo" />
                </div>
            </div>
            <div id="meio" style="background-color: White;">
                <div id="meioconteudo">
                    <table width="100%">
                        <tr>
                            <td valign="top">
                                <div style="padding: 10px">
                                    <!-- Widgets Notícias Agrícolas - www.noticiasagricolas.com.br/widgets -->
                                    <script type="text/javascript" src="https://www.noticiasagricolas.com.br/widgets/noticias?subsecao=2,3,162,6,7,8,10,64,80,85,146,5,4,11,160,12,156,40,158,60,13,163,97,14,95,205,15,1,155,84,28,149,26,69,90,62,27,92,148,154,32,67,101,102,103,105,207,106,107,108,109,147,206,110,111,112,113,159,114,157,115,116,210,68,164,117,118,119,204,211,120,121,203,166&largura=auto&altura=auto&fonte=Arial%2C%20Helvetica%2C%20sans-serif&tamanho=10pt&cortexto=333333&corlink=006666&qtd=13&output=js"></script>
                                </div>
                            </td>
                            <td valign="top" style="width: 10px;">&nbsp;
                            </td>
                            <td valign="top" style="float: right;">
                                <asp:UpdatePanel ID="updpnlBancoServidor" runat="server">
                                    <ContentTemplate>
                                        <div style="width: 275px; height: 275px; background: none repeat scroll 0 0 #F5F5F5; border: 1px solid #E5E5E5; margin: 12px 0 0; padding: 20px 25px 15px;">
                                            <h2 class="iconPC" style="font-size: 15pt; height: 16px; line-height: 16px; margin: 0 0 1.2em; 
                                                position: relative;">Login <strong style="display: inline-block; height: 19px; 
                                                position: absolute; right: 0; top: 1px; width: 52px;"></strong>
                                            </h2>
                                            <asp:Panel ID="loginMenu" runat="server">
                                                <div style="text-align: center;">
                                                    <asp:Label ID="lblErro" runat="server" ForeColor="Red" Font-Bold="true" Visible="false" />
                                                </div>
                                                <br />
                                                <label style="display: block; margin: 0 0 1.5em;">
                                                    <strong style="-moz-user-select: none; display: block; font-weight: bold; margin: 0 0 0.5em; font-size: 12pt">Nome de usuário</strong>
                                                    <asp:TextBox ID="txtUsuario" runat="server" Style="font-size: 12pt; height: 32px; width: 100%;"
                                                        autofocus />
                                                </label>
                                                <label style="display: block; margin: 0 0 1.5em;">
                                                    <strong style="-moz-user-select: none; display: block; font-weight: bold; margin: 0 0 0.5em; font-size: 12pt">Senha</strong>
                                                    <asp:TextBox ID="txtSenha" runat="server" Style="font-size: 12pt; height: 32px; width: 100%;"
                                                        TextMode="Password" />
                                                </label>
                                                <label style="display: block; margin: 0 0 1.5em;">
                                                    <asp:DropDownList ID="cmbBancoServidor" runat="server" Visible="False" Style="font-size: 8pt; height: 32px; width: 289px;" />
                                                </label>
                                                <asp:Button ID="btnEntrar" runat="server" Text="Entrar" OnClick="btnEntrar_Click"
                                                    CssClass="botaologin" Height="38px" />
                                                 <asp:Button ID="btnRecuperarSenha" runat="server" style="border:none;background-color:transparent;color:#4D90FE;font-size:14px;" Text="Esqueci minha senha" 
                                                     OnClick="btnRecuperarSenha_Click" Height="38px" />
                                            </asp:Panel>
                                            <asp:Panel ID="recuperarSenhaMenu" runat="server" Visible="False">
                                                <label style="display: block; margin: 0 0 1.5em;">
                                                    <strong style="-moz-user-select: none; display: block; font-weight: bold; margin: 0 0 0.5em; font-size: 12pt">Usuário de Recuperação</strong>
                                                    <asp:TextBox ID="txtUsuarioRecuperacao" runat="server" Style="font-size: 12pt; height: 32px; width: 100%;"
                                                        autofocus />
                                                </label>
                                                <label style="display: block; margin: 0 0 1.5em;">
                                                    <asp:DropDownList ID="cmbBancoPorUsuario" runat="server" Style="font-size: 8pt; height: 32px; width: 289px;" />
                                                </label>
                                                <asp:Label ID="lblMesagemEnvioEmail" runat="server" Visible="false"/>
                                                <asp:Button ID="btnVoltar" runat="server" Text="Voltar" OnClick="btnVoltar_Click"
                                                    CssClass="botaologin" Height="38px" />
                                                <asp:Button ID="btnEnviarEmail" runat="server" Text="Enviar" OnClick="btnEnviarEmail_Click"
                                                    CssClass="botaologin" Height="38px" />                         
                                            </asp:Panel>
                                        </div>
                                    </ContentTemplate>
                                </asp:UpdatePanel>
                            </td>
                        </tr>
                    </table>
                </div>
            </div>
        </div>
        <div id="rodape">
            <div id="rodapemeio">
                Versão
            <asp:Label ID="lblVersao" runat="server" />
                &nbsp;<span class="copyright">&copy;
                <%= DateTime.Now.Year.ToString() %>
                NGS Sistemas Ltda.</span>
            </div>
        </div>
    </form>
</body>
</html>
