<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="PosicaoDeRomaneios.aspx.vb" Inherits="NGS.Web.UI.PosicaoDeRomaneios" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngPosicaoDeRomaneios" runat="server" AsyncPostBackTimeout="900" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updnlPosicaoDeRomaneios" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Posição de Romaneios
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconRelatorio" runat="server">
                            <asp:LinkButton ID="lnkExcel" Text="Excel" runat="server" />
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li class="iconAjuda" runat="server">
                            <asp:LinkButton ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUnidadeNegocio" runat="server" Width="630px" AutoPostBack="True"
                        OnSelectedIndexChanged="cmbUnidadeNegocio_SelectedIndexChanged" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="630px" />
                </div>
            </div>
            <div class="row">
                <uc:SelecaoProduto ID="ucSelecaoProduto" runat="server" />
            </div>
            <div class="row">
                <div class="collbl">
                    Notas:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="RadEntradas" runat="server" Text=" Entradas " Checked="True" GroupName="Notas" data-ToolTip="default" ToolTip="Selecionar notas entrada ou saida." />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="RadSaidas" runat="server" Text=" Saídas " GroupName="Notas" data-ToolTip="default" ToolTip="Selecionar notas entrada ou saida." />
                </div>
                <div class="collbl">
                    Período Inicial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicial" runat="server" CssClass="calendario" data-ToolTip="default" ToolTip="Período inicial da consulta." />
                </div>
                <div class="collbl">
                    Período Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" runat="server" CssClass="calendario" data-ToolTip="default" ToolTip="Período final da consulta." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
