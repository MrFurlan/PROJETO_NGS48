<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="NotaFiscalAjustes.aspx.vb"
    Inherits="NGS.Web.UI.NotaFiscalAjustes" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <link href="Styles/menucentral.css" rel="stylesheet" type="text/css" />
    <title>Livro Registro de Apuração do ICMS - Ajustes Outros Débitos/Créditos</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:ScriptManager ID="scmngNotaFiscalAjustes" runat="server" />
        <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    </div>
    <asp:UpdatePanel ID="updpnlNotaFiscalAjustes" runat="server">
        <ContentTemplate>
            <table style="width: 480px">
                <tr>
                    <td colspan="3" class="titulotabela">
                        <label>
                            Nota Fiscal - Revisão
                        </label>
                    </td>
                </tr>
                <tr>
                    <td style="color: red" class="rotulo">
                        Quantidade.:
                    </td>
                    <td>
                        <asp:TextBox ID="txtPesoFiscal" runat="server" MaxLength="9" Width="190px"/>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td style="color: red" class="rotulo">
                        Unitário.:
                    </td>
                    <td rowspan="1">
                        <asp:TextBox ID="txtUnitario" runat="server" Width="190px"/>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td class="rotulo" style="color: red">
                        Total
                    </td>
                    <td>
                        <asp:TextBox ID="txtValorTotal" runat="server" Width="190px"/>
                    </td>
                    <td align="right">
                        <asp:Button ID="cmdCalcular" runat="server" CssClass="botao" OnClick="cmdCalcular_Click"
                            Text="Calcular" />
                    </td>
                </tr>
                <tr>
                    <td class="rotulo">
                        Base Icms.:
                    </td>
                    <td>
                        <asp:TextBox ID="txtBaseIcms" runat="server" Width="190px"/>
                    </td>
                    <td align="right">
                        <asp:Button ID="cmdLimpar" runat="server" OnClick="cmdLimpar_Click" Text="Limpar"
                            CssClass="botao" />
                    </td>
                </tr>
                <tr>
                    <td class="rotulo">
                        Valor Icms.:
                    </td>
                    <td>
                        <asp:TextBox ID="txtValorIcms" runat="server" Width="190px"/>
                    </td>
                    <td align="right">
                        <asp:Button ID="cmdSair" runat="server" CssClass="botao" OnClick="cmdSair_Click"
                            Text="Sair" />
                    </td>
                </tr>
                <tr>
                    <td colspan="3">
                    </td>
                </tr>
            </table>
        </ContentTemplate>
    </asp:UpdatePanel>
    </form>
</body>
</html>
