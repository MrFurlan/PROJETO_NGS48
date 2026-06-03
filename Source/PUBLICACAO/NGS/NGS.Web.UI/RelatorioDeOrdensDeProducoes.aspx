<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master" CodeBehind="RelatorioDeOrdensDeProducoes.aspx.vb" Inherits="NGS.Web.UI.RelatorioDeOrdensDeProducoes" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        #meioconteudo {
            width: 1280px !important;
        }
    </style>
    <script type="text/javascript">
        function pageLoadRelatorioDeOrdensDeProducoes() {
            $("#txtLote").setMask("lote-producao");
        }

        $(document).ready(function () {

            pageLoadRelatorioDeOrdensDeProducoes();
            var prmItemProducao = Sys.WebForms.PageRequestManager.getInstance();
            prmItemProducao.add_endRequest(pageLoadRelatorioDeOrdensDeProducoes);

        });

    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngOrdemDeProducao" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />

    <asp:UpdatePanel ID="updpnlOrdemDeProducao" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Relatório de Ordens de produções
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconRelatorio rel" runat="server"><a>Relatório</a>
                            <ul>
                                <li class="iconPdf">
                                    <asp:LinkButton ID="lnkRelatorioPDF" runat="server" Text="PDF" />
                                </li>
                                <li class="iconExcel">
                                    <asp:LinkButton ID="lnkRelatorioExcel" runat="server" Text="Excel" />
                                </li>
                            </ul>
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
                <div class="collbl">
                    Data inicial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicial" CssClass="calendario" runat="server" data-ToolTip="default"
                        ToolTip="Data inicial para o Relatório." />
                </div>
                <div class="collbl">
                    Data final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" CssClass="calendario" runat="server" data-ToolTip="default"
                        ToolTip="Data final para o Relatório." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Tipo:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rbDtMov" runat="server" Checked="True" GroupName="gData"
                        Text="Movimento" data-ToolTip="default" ToolTip="Por data de Movimento." />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rbDtVen" runat="server" GroupName="gData"
                        Text="Validade" data-ToolTip="default" ToolTip="Por data de Validade." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Ordem:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rbOrdem" runat="server" Checked="True" GroupName="gOrdem"
                        Text="Ordem" data-ToolTip="default" ToolTip="Ordena por Ordem." />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rbOPrd" runat="server" GroupName="gOrdem"
                        Text="Produto" data-ToolTip="default" ToolTip="Ordena por Produto." />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rbOPrdNome" runat="server" GroupName="gOrdem"
                        Text="Nome do Produto" data-ToolTip="default" ToolTip="Ordena por Nome do Produto." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Tipo:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rbAberta" runat="server" Checked="True" GroupName="gTipo"
                        Text="Aberta(s)" data-ToolTip="default" ToolTip="Ordens Aberta(s)." />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rbEncerrada" runat="server" GroupName="gTipo"
                        Text="Encerrada(s)" data-ToolTip="default" ToolTip="Ordens Encerrada(s)." />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rbCancelada" runat="server" GroupName="gTipo"
                        Text="Cancelada(s)" data-ToolTip="default" ToolTip="Ordens Cancelada(s)." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
