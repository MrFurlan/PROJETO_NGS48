<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="SpedPisCofins.aspx.vb" Inherits="NGS.Web.UI.SpedPisCofins" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <script type="text/javascript">
        function downloadArquivo() {
            $("#MainContent_imdDownload").click();
        }
    </script>
    <style type="text/css">
        .collbl {
            width: 170px;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngSpedPisCofins" runat="server" AsyncPostBackTimeout="50000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlSpedPisCofins" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                EFD Contribuições
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconRelatorio" runat="server">
                            <asp:LinkButton ID="lnkProcessar" Text="Processar" runat="server" />
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
                    <asp:DropDownList ID="ddlUnidade" runat="server" AutoPostBack="True" OnSelectedIndexChanged="DdlUnidade_SelectedIndexChanged"
                        Width="634px" />
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
                    Periodo:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicial" runat="server" Width="96px" CssClass="calendario" data-ToolTip="default" ToolTip="Informar o período de apuração." />
                    &nbsp;a&nbsp;
                    <asp:TextBox ID="txtDataFinal" runat="server" Width="96px" CssClass="calendario" data-ToolTip="default" ToolTip="Informar o período de apuração." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Tipo:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rbConsolidado" runat="server" Checked="True" Text="Consolidado - Blocos (C180,181,185/C190,C191,C195)" GroupName="TipoArquivo" data-ToolTip="default" ToolTip="Gera blocos (C180,181,185/C190,C191,C195)" />
                    <asp:RadioButton ID="rbAnalitico" runat="server" Text="Analítico - Blocos (C100,C170)" GroupName="TipoArquivo" data-ToolTip="default" ToolTip="Gera blocos (C100,C170)" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Incidencia:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cboIncidencia" runat="server" Width="634px">
                        <asp:ListItem Value="1">1 - Incidência Não-Cumulativa</asp:ListItem>
                        <asp:ListItem Value="2">2 - Incidência Cumulativa</asp:ListItem>
                        <asp:ListItem Value="3">3 - Incidência Cumulativa e Não-Cumulativa</asp:ListItem>
                    </asp:DropDownList>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Tipo de Escrituracao:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cboTipoEscrituracao" runat="server" Width="634px">
                        <asp:ListItem Value="0">0 - Original</asp:ListItem>
                        <asp:ListItem Value="1">1 - Retificadora</asp:ListItem>
                    </asp:DropDownList>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Indicador de Situacao Especial:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlIndSituacaoEsp" runat="server" Width="634px">
                        <asp:ListItem Value="">     Normal</asp:ListItem>
                        <asp:ListItem Value="0">0 - Abertura</asp:ListItem>
                        <asp:ListItem Value="1">1 - Cis&#227;o</asp:ListItem>
                        <asp:ListItem Value="2">2 - Fus&#227;o</asp:ListItem>
                        <asp:ListItem Value="3">3 - Incorpora&#231;&#227;o</asp:ListItem>
                        <asp:ListItem Value="4">4 - Extinç&#227;o</asp:ListItem>
                    </asp:DropDownList>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Indicador da Natureza:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlIndicadorDeNatureza" runat="server" Width="634px">
                        <asp:ListItem Value="00">00 - Sociedade empres&#225;ria em geral</asp:ListItem>
                        <asp:ListItem Value="01">01 - Sociedadecooperativa</asp:ListItem>
                        <asp:ListItem Value="02">02 - Entidade sujeita ao PIS/Pasep Exclusivamente com base na Folha de Sal&#225;rios</asp:ListItem>
                    </asp:DropDownList>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Indicador de Atividade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlIndicadorAtividade" runat="server" Width="634px">
                        <asp:ListItem Value="0">0 - Industrial ou equiparado a industrial</asp:ListItem>
                        <asp:ListItem Value="1">1 - Prestador de Servi&#231;os</asp:ListItem>
                        <asp:ListItem Value="2">2 - Atividade de com&#233;rcio</asp:ListItem>
                        <asp:ListItem Value="3">3 - Atividade financeira</asp:ListItem>
                        <asp:ListItem Value="4">4 - Atividade imobili&#225;ria</asp:ListItem>
                        <asp:ListItem Value="9">9 - Outros</asp:ListItem>
                    </asp:DropDownList>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Numero Recibo Anterior:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txt_NUM_ORD" runat="server" Width="630px" data-ToolTip="default" ToolTip="Informar o número do recibo da entrega anterior." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Arquivo de Saida:
                </div>
                <div class="coltxt">
                    <asp:UpdatePanel ID="updpnlArquivoDeSaida" runat="server">
                        <ContentTemplate>
                            <asp:TextBox ID="txtArquivoDeSaida" Enabled="false" runat="server" Width="630px" data-ToolTip="default" ToolTip="Nome do arquivo da transmissão." />
                            <asp:ImageButton ID="imdDownload" runat="server" ImageAlign="AbsMiddle" ImageUrl="~/Images/download32x32.png" Style="margin-top: 0;"
                                Height="22px" Width="22px" OnClick="imdDownload_Click" ToolTip="Baixar Arquivo" />
                        </ContentTemplate>
                        <Triggers>
                            <asp:PostBackTrigger ControlID="imdDownload" />
                        </Triggers>
                    </asp:UpdatePanel>
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
