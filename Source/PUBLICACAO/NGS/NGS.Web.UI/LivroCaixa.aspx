<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="LivroCaixa.aspx.vb" Inherits="NGS.Web.UI.LivroCaixa" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngLivroCaixa" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlLivroCaixa" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Livro Caixa
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconRelatorio" ID="lnkRelatorio" Text="Relatório" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconLimpar" ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlUnidadeDeNegocio" runat="server" AutoPostBack="True" OnSelectedIndexChanged="DdlUnidadeDeNegocio_SelectedIndexChanged"
                        Width="618px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlEmpresa" runat="server" Width="618px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Conta:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoConta" runat="server" />
                    <asp:TextBox ID="txtConta" runat="server" Width="584px" Enabled="False" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnCodigoConta" runat="server" Text=">" CssClass="btn" data-ToolTip="default"
                        ToolTip="Informar a conta desejada." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtData" runat="server" Width="106px" CssClass="calendario" data-ToolTip="default"
                        ToolTip="Informar a data da pesquisa." />
                </div>
                <div class="collbl">
                    Número:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNumero" CssClass="txtNumerico" runat="server" Width="112px" AutoPostBack="True" OnTextChanged="txtNumero_TextChanged" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Saldo Anterior:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtSaldoAnterior" CssClass="txtDecimal" runat="server" Enabled="False" />
                </div>
                <div class="collbl" style="margin-left: 27px;">
                    Dinheiro:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDinheiro" CssClass="txtDecimal" runat="server" Width="112px"
                        TabIndex="14" AutoPostBack="True" OnTextChanged="txtDinheiro_TextChanged" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Entradas do Dia:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtEntradasDia" CssClass="txtDecimal" runat="server" Enabled="False" />
                </div>
                <div class="collbl" style="margin-left: 27px;">
                    Cheques:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCheques" CssClass="txtDecimal" runat="server" Width="112px" AutoPostBack="True"
                        TabIndex="14" OnTextChanged="txtCheques_TextChanged" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Saídas do Dia:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtSaidasDia" CssClass="txtDecimal" runat="server" Enabled="False" />
                </div>
                <div class="collbl" style="margin-left: 27px;">
                    Valores:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtVales" CssClass="txtDecimal" runat="server" Width="112px" AutoPostBack="True"
                        TabIndex="14" OnTextChanged="txtVales_TextChanged" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Saldo Atual:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtSaldoAtual" CssClass="txtDecimal" runat="server" Enabled="False" />
                </div>
                <div class="collbl" style="margin-left: 27px;">
                    Selos:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtSelos" runat="server" Width="112px" CssClass="txtDecimal" AutoPostBack="True"
                        OnTextChanged="txtSelos_TextChanged" TabIndex="14" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Total:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtTotal" CssClass="txtDecimal" runat="server" Enabled="False" />
                </div>
                <div class="collbl" style="margin-left: 27px;">
                    Outros:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtOutros" CssClass="txtDecimal" runat="server" Width="112px" AutoPostBack="True"
                        OnTextChanged="txtOutros_TextChanged" TabIndex="14" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaPlanoDeContas ID="ucConsultaPlanoDeContas" runat="server" />
</asp:Content>
