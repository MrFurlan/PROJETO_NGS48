<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SenhaNova.aspx.vb" Inherits="NGS.Web.UI.SenhaNova" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head runat="server">
    <link href="Styles/menucentral.css" rel="stylesheet" type="text/css" />
    <title>NGS Soluções Ltda</title>
    <link rel="shortcut icon" href="Images/user_group.ico" />
    <link href="Styles/redmond/jquery-ui.custom.css" rel="stylesheet" type="text/css" />
    <link href="Styles/jquery.loadmask.css" rel="stylesheet" type="text/css" />
    <script src="Scripts/jquery.min.js" type="text/javascript"></script>
    <script src="Scripts/jquery-ui.custom.min.js" type="text/javascript"></script>
    <script type="text/javascript">
        $(document).ready(function () {
            $("input[type='button'].botao").button();
            $("input[type='submit'].botao").button();
        });
    </script>
    <style type="text/css">
        #box {
            height: 200px;
            left: 50%;
            margin: -100px 0 0 -200px;
            position: absolute;
            top: 50%;
            width: 400px;
            text-align: center;
            vertical-align: middle;
        }

        .subtitulo {
            background-color: #5D7B9D;
            color: #FFFFFF;
            height: 24px;
            font-weight: bolder;
            vertical-align: middle;
            text-align: center;
            text-transform: uppercase;
        }
    </style>
</head>
<body>
    <form runat="server">
        <asp:ScriptManager ID="scrmSenhaNova" runat="server" />
        <div id="box">
            <table class="bordasimples" style="width: 100%;">
                <tr>
                    <td class="subtitulo" style="text-align: center;">CADASTRO DE SENHA
                    </td>
                </tr>
                <tr>
                    <td style="height: 95px; text-align: center;">
                        <table style="border: 0px none; text-align: center;">
                            <tr>
                                <td style="text-align: right;">Senha:
                                </td>
                                <td>
                                    <asp:TextBox ID="txtSenha" runat="server" CssClass="cForm required" TextMode="Password"
                                        MaxLength="50" Style="padding: 5px; font-size: 14px;" Width="200px" />
                                </td>
                            </tr>
                            <tr>
                                <td style="text-align: right;">Confirmação:
                                </td>
                                <td>
                                    <asp:TextBox ID="txtConfirmacao" runat="server" CssClass="cForm required" TextMode="Password"
                                        MaxLength="50" Style="padding: 5px; font-size: 14px;" Width="200px" />
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr>
                    <td style="text-align: center;">
                        <hr style="width: 90%;" />
                    </td>
                </tr>
                <tr>
                    <td style="padding-bottom: 10px; padding-top: 5px; text-align: center; vertical-align: middle;">
                        <asp:Button ID="btnConfirmar" runat="server" CssClass="botao" Text="CONFIRMAR" Style="width: 100px;"
                            OnClick="btnConfirmar_Click" />
                        <asp:Button ID="btnCancelar" runat="server" CssClass="botao" Text="CANCELAR" Style="width: 100px;"
                            OnClick="btnCancelar_Click" />
                    </td>
                </tr>
            </table>
        </div>
    </form>
</body>

</html>
