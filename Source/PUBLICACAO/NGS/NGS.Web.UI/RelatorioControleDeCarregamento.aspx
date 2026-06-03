<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="RelatorioControleDeCarregamento.aspx.vb" Inherits="NGS.Web.UI.RelatorioControleDeCarregamento" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <script type="text/javascript">
        function pageLoadRCC() {
            $("input[type='checkbox']", "#MainContent_ucSelecaoProduto_gridProduto").each(function () {
                if ($(this).attr("checked") == "checked" && $(this).parent().next().text() == '101010003') {
                    $('#<%=INTACTA.ClientID%>').show();
                    $('#<%=hIntacta.ClientID%>').val(true);
                } else {
                    $('#<%=INTACTA.ClientID%>').hide();
                    $('#<%=hIntacta.ClientID%>').val(false);
                }
            });
        }

        $(document).ready(function () {
            pageLoadRCC();
            var prmRCC = Sys.WebForms.PageRequestManager.getInstance();
            prmRCC.add_endRequest(pageLoadRCC);
        });
    </script>
    <style type="text/css">
        .collbl {
            width: 165px;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngRelatorioControleDeCarregamento" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" AsyncPostBackTimeout="10000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlRelatorioControleDeCarregamento" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <asp:HiddenField ID="hIntacta" runat="server" />
            <div class="titulodiv">
                Controle De Carregamento/Troca De Nota
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconRelatorio" runat="server">
                            <asp:LinkButton ID="lnkRelatorio" runat="server">
                                        <span>Relatório</span>
                            </asp:LinkButton>
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" runat="server">
                                        <span>Limpar</span>
                            </asp:LinkButton>
                        </li>
                        <li class="iconAjuda" runat="server">
                            <asp:LinkButton ID="lnkAjuda" runat="server">
                                        <span>Ajuda</span>
                            </asp:LinkButton>
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbEmpresa" runat="server" Width="626px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Relatório:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdOrigemDestino" Text="Origem / Destino" GroupName="Rel" Checked="true"
                        runat="server" data-ToolTip="default" ToolTip="Selecionar a opção da consulta." />
                    <asp:RadioButton ID="rdDestinoOrigem" Text="Destino / Origem" GroupName="Rel" runat="server"
                        data-ToolTip="default" ToolTip="Selecionar a opção da consulta." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkConsolidarClienteOrigem" Text="Consolidar Cliente Origem:" runat="server" />
                    <asp:HiddenField ID="txtCodigoCliOrigem" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtClienteOrigem" runat="server" Width="587px" Enabled="False" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="cmdBuscaCliOrigem" OnClick="cmdBuscaCliOrigem_Click" runat="server"
                        UseSubmitBehavior="False" CssClass="btn" Text=">" data-ToolTip="default" ToolTip="Unificar as informações por cliente de origem." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkConsolidarClienteDestino" runat="server" Text="Consolidar Cliente Destino:" />
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtClienteDestino" runat="server" Width="587px" Enabled="False" />
                    <asp:HiddenField ID="txtCodigoCliDestino" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="cmdBuscaCliDestino" OnClick="cmdBuscaCliDestino_Click" runat="server"
                        UseSubmitBehavior="False" CssClass="btn" Text=">" data-ToolTip="default" ToolTip="Unificar as informações por cliente de destino." />
                </div>
            </div>
            <div class="row">
                <div class="coltxt" style="width: 100%; margin-bottom: 4px;">
                    <uc:SelecaoProduto ID="ucSelecaoProduto" runat="server" />
                </div>
            </div>
            <div id="INTACTA" runat="server" class="row" style="display: none;">
                <div class="collbl">
                    Intacta:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdITodos" runat="server" Text="Todos" data-ToolTip="default"
                        ToolTip="Listar Todos" Checked="True" GroupName="intacta" />
                    &nbsp; &nbsp;
                    <asp:RadioButton ID="rdISim" runat="server" Text="Apenas Intacta" data-ToolTip="default"
                        ToolTip="Lista apenas Laudos Intacta" GroupName="intacta" />
                    &nbsp; &nbsp;
                    <asp:RadioButton ID="rdPositivo" runat="server" Text="Apenas Teste Positivo" data-ToolTip="default"
                        ToolTip="Lista apenas Laudos Teste Positivo" GroupName="intacta" />
                    &nbsp; &nbsp;
                    <asp:RadioButton ID="rdPSim" runat="server" Text="Intacta e Teste Positivo" data-ToolTip="default"
                        ToolTip="Lista apenas Laudos Intacta" GroupName="intacta" />
                    &nbsp; &nbsp;
                    <asp:RadioButton ID="rdINao" runat="server" Text="Sem Intacta" data-ToolTip="default"
                        ToolTip="Lista apenas Laudos sem Intacta" GroupName="intacta" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Operação:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbOperacao" runat="server" Width="626px" AutoPostBack="True"
                        OnSelectedIndexChanged="cmbOperacao_SelectedIndexChanged" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Sub Operação:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbSubOperacao" runat="server" Width="626px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data Inicial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicial" CssClass="calendario" runat="server" Width="112px"
                        data-ToolTip="default" ToolTip="Informar o período inicial de consulta." />
                </div>
                <div class="collbl" style="width: 120px;">
                    Data Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" CssClass="calendario" runat="server" Width="112px"
                        data-ToolTip="default" ToolTip="Informar o período final de consulta." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
