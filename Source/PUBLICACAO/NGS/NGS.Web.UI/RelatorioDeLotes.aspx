<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="RelatorioDeLotes.aspx.vb" Inherits="NGS.Web.UI.RelatorioDeLotes" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngRelatorioDeLotes" runat="server" AsyncPostBackTimeout="900" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlRelatorioDeLotes" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Relatório de Lotes
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconRelatorio rel" ID="lnkRelatorio" runat="server"
                                Text="Relatório" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconLimpar" ID="lnkLimpar" runat="server"
                                Text="Limpar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" runat="server"
                                Text="Ajuda" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUnidade" runat="server" Width="634px" OnSelectedIndexChanged="DdlUnidade_SelectedIndexChanged"
                        AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlEmpresa" runat="server" Width="634px" />
                </div>
            </div>
            <div class="row">
                <div class="coltxt lg">
                    <uc:SelecaoProduto ID="ucSelecaoProduto" runat="server" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkPeriodo" runat="server" AutoPostBack="True" Font-Bold="True"
                        Text="Usar Data" />
                </div>
                <div class="coltxt">
                    <asp:Panel ID="pnlDataMovimento" runat="server" Visible="False" HorizontalAlign="Left">
                        &nbsp;<asp:TextBox ID="txtDataInicial" runat="server" CssClass="calendario" Width="112px"
                            data-ToolTip="default" ToolTip="Selecionar o período desejado." />
                        &nbsp;a
                        <asp:TextBox ID="txtDataFinal" runat="server" CssClass="calendario" Width="112px"
                            data-ToolTip="default" ToolTip="Selecionar o período desejado." />
                    </asp:Panel>
                </div>
            </div>
            <div class="row" id="tipoLista" runat="server">
                <div class="collbl">
                    Tipo:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdMovimento" runat="server" Text="Movimento" Checked="True" GroupName="Ordem" data-ToolTip="default" ToolTip="Filtra o relatório pelo movimento." />
                    <asp:RadioButton ID="rdValidade" runat="server" Text="Validade" GroupName="Ordem" data-ToolTip="default" ToolTip="Filtra o relatório pela Validade." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
