<%@ Page Title="" Language="vb" AutoEventWireup="false" CodeBehind="CPRxFazenda.aspx.vb"
    Inherits="NGS.Web.UI.CPRxFazenda" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>CPR x Fazenda</title>
    <link href="Styles/menucentral.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form runat="server">
    <asp:ScriptManager ID="scrmngCPRxFazenda" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlCPRxFazenda" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <table style="width: 100%; border: 0px none;">
                <tr>
                    <td class="titulotabela">
                        CPR x Fazenda
                    </td>
                </tr>
                <tr>
                    <td style="vertical-align: top;">
                        Cliente.:<asp:TextBox ID="txtFazenda" runat="server" Enabled="False" Width="420px"/><asp:Button
                            ID="btnFazenda" runat="server" Text=">" UseSubmitBehavior="False" OnClick="btnFazenda_Click" /><asp:HiddenField
                                ID="txtCodigoFazenda" runat="server" />
                    </td>
                </tr>
                <tr>
                    <td>
                        <asp:GridView ID="gridMatricula" runat="server" CellPadding="4" ForeColor="#333333"
                            GridLines="None" Width="450px" AutoGenerateColumns="False">
                            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                            <EditRowStyle BackColor="#999999" />
                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                            <Columns>
                                <asp:BoundField DataField="Matricula_Id" HeaderText="Matricula" />
                                <asp:BoundField DataField="Area" HeaderText="Area" />
                                <asp:TemplateField HeaderText="Area CPR">
                                    <ItemTemplate>
                                        <asp:TextBox ID="txtArea" runat="server" onkeyup="validar_valor(this, 15);" value="0,00"
                                            Text='<%# eval("AreaCpr") %>'/>
                                    </ItemTemplate>
                                </asp:TemplateField>
                            </Columns>
                        </asp:GridView>
                    </td>
                </tr>
                <tr>
                    <td style="height: 26px; text-align: right;">
                        <asp:Button ID="btnConfirmar" runat="server" Text="Confirmar" CssClass="botao" OnClick="btnConfirmar_Click" />
                        <asp:Button ID="btnSair" runat="server" CssClass="botao" Text="Sair" OnClick="btnSair_Click" />&nbsp;
                    </td>
                </tr>
            </table>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    </form>
</body>
</html>
