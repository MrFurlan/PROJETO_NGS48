<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="CalculaJurosAdiantamentos.aspx.vb" Inherits="NGS.Web.UI.CalculaJurosAdiantamentos" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <link href="Styles/menucentral.css" rel="stylesheet" type="text/css" />
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngCalculaJurosAdiantamentos" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlCalculaJurosAdiantamentos" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <table style="width: 100%; border: 0px none">
                <tr>
                    <td class="titulotabela" colspan="3">
                        <label>
                            Cálculo de Juros
                        </label>
                    </td>
                </tr>
                <tr>
                    <td style="width: 100px">
                        <div class="headerGray">
                            <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Empresa:</span>
                            <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                        </div>
                    </td>
                    <td>
                        <asp:TextBox ID="txtCnpjDaEmpresa" runat="server" CssClass="texto" Width="120px"
                            Enabled="False"></asp:TextBox>
                        <asp:TextBox ID="txtEndEmpresa" runat="server" CssClass="texto" Width="50px" Enabled="False"></asp:TextBox>
                        <asp:TextBox ID="txtNomeDaEmpresa" runat="server" CssClass="texto" Width="300px"
                            Enabled="False"></asp:TextBox>
                        <asp:ImageButton ID="ImgEmpresa" OnClick="ImgEmpresa_Click" runat="server" Width="21px"
                            ImageAlign="Middle" ToolTip="Buscar Empresas" ImageUrl="~/images/Lupa.jpg"></asp:ImageButton>
                    </td>
                    <td>
                        <asp:Button ID="btnCalcular" OnClick="btnNovo_Click" runat="server" CssClass="botao"
                            Text="Calcular"></asp:Button>
                    </td>
                </tr>
                <tr>
                    <td>
                        <div class="headerGray">
                            <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Cliente:</span>
                            <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                        </div>
                    </td>
                    <td>
                        <asp:TextBox ID="txtCnpjDoCliente" runat="server" CssClass="texto" Width="120px"
                            Enabled="False"></asp:TextBox>
                        <asp:TextBox ID="txtEndDoCliente" runat="server" CssClass="texto" Width="50px" Enabled="False"></asp:TextBox>
                        <asp:TextBox ID="txtNomeDoCliente" runat="server" CssClass="texto" Width="300px"
                            Enabled="False"></asp:TextBox>
                        <asp:ImageButton ID="ImgClientes" OnClick="ImgClientes_Click" runat="server" Width="21px"
                            ImageAlign="Middle" ToolTip="Buscar Clientes" ImageUrl="~/images/Lupa.jpg" Height="21px">
                        </asp:ImageButton>
                    </td>
                    <td>
                        <input id="btnSair" class="botao" onclick="MenuDeAcesso('Adiantamento');" type="button"
                            value="Sair" />
                    </td>
                </tr>
                <tr>
                    <td>
                        <div class="headerGray">
                            <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Data:</span>
                            <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                        </div>
                    </td>
                    <td colspan="2">
                        <asp:TextBox ID="txtMovimento" CssClass="calendario" runat="server"></asp:TextBox>
                    </td>
                </tr>
            </table>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaEmpresas ID="ucConsultaEmpresas" runat="server" />
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
