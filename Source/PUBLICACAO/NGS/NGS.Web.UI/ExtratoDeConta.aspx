<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="ExtratoDeConta.aspx.vb" Inherits="NGS.Web.UI.ExtratoDeConta" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngExtratoDeConta" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlExtratoDeConta" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Extrato de Contas
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconRelatorio rel" runat="server"><a>Relatório</a>
                            <ul>
                                <li runat="server">
                                    <asp:LinkButton class="iconPdf" ID="lnkPdf" runat="server" Text="Pdf" />
                                </li>
                                <li runat="server">
                                    <asp:LinkButton class="iconExcel" ID="lnkExcel" runat="server" Text="Excel" />
                                </li>
                            </ul>
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
                    <asp:DropDownList ID="ddlUnidade" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlUnidade_SelectedIndexChanged"
                        Width="631px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                    <asp:CheckBox ID="chkConsolidarEmpresa" ToolTip="Unificar o CNPJs da empresa." CssClass="multiselect" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="631px" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Conta:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="hdfConta" runat="server" />
                    <asp:TextBox ID="txtConta" runat="server" Width="591px" Enabled="false" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnConsultaConta" runat="server" Text=">" CssClass="btn" data-ToolTip="default"
                        ToolTip="Selecionar a conta desejada." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="hdfCliente" runat="server" />
                    <asp:TextBox ID="txtCliente" runat="server" Enabled="false" Width="591px" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnCliente" runat="server" Text=">" CssClass="btn" data-ToolTip="default"
                        ToolTip="Selecionar o cliente desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Centro de Custo:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlCentroDeCusto" runat="server" Width="631px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Grupo:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlGrupo" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlGrupo_SelectedIndexChanged1"
                        Width="631px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Produto:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlProduto" runat="server" Width="631px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicial" CssClass="calendario" runat="server" Width="112px"
                        data-ToolTip="default" ToolTip="Período de pesquisa." />
                    &nbsp;<asp:TextBox ID="txtDataFinal" CssClass="calendario" runat="server" Width="120px"
                        data-ToolTip="default" ToolTip="Período de pesquisa." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Filtros:
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="ckPorCliente" Text="Por Cliente" runat="server" AutoPostBack="True"
                        data-ToolTip="default" ToolTip="Filtrar por cliente ou histórico." />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkHistorico" runat="server" Text="Ordenar por Histórico" data-ToolTip="default"
                        ToolTip="Filtrar por cliente ou histórico." />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkContasDoPeriodo" runat="server" AutoPostBack="True" Visible="false" Text="Todas as Contas do Período selecionado" data-ToolTip="default"
                        ToolTip="Todas as Contas do Período selecionado." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaPlanoDeContas ID="ucConsultaPlanoDeContas" runat="server" />
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
