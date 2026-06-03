<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="FCont.aspx.vb" Inherits="NGS.Web.UI.FCont" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        .collbl {
            width: 165px;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngFCont" runat="server" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlFCont" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Sped FCont
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
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="618px" AutoPostBack="True"
                        OnSelectedIndexChanged="ddlEmpresa_SelectedIndexChanged" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data Inicial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicial" CssClass="calendario" runat="server" Width="86px"
                        data-ToolTip="default" ToolTip="Período inicial da pesquisa." />
                </div>
                <div class="collbl" style="width: 100px">
                    Data Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" CssClass="calendario" runat="server" Width="86px"
                        data-ToolTip="default" ToolTip="Período final da pesquisa." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Conta De Resultado:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtContaDeResultado" runat="server" Width="86px" data-ToolTip="default"
                        ToolTip="Informar uma conta de resultado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Tipo De Escrituração:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlTipoDeEscrituracao" runat="server" Width="618px">
                        <asp:ListItem></asp:ListItem>
                        <asp:ListItem>0 - Original</asp:ListItem>
                        <asp:ListItem>1 - Retificadora</asp:ListItem>
                    </asp:DropDownList>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Número do Recibo Anterior:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNumeroDoReciboAnterior" runat="server" Width="86px" data-ToolTip="default"
                        ToolTip="Informar o número do ultimo recibo gerado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Situação Sdo da Escrituração:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlSituacaoSdoDaEscrituracao" runat="server" Width="618px">
                        <asp:ListItem></asp:ListItem>
                        <asp:ListItem>N - N&#227;o Recuperado</asp:ListItem>
                        <asp:ListItem>R - Recuperado</asp:ListItem>
                        <asp:ListItem>E - Editado</asp:ListItem>
                        <asp:ListItem>I - Importado</asp:ListItem>
                    </asp:DropDownList>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Forma de Apuração:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlFormaDeApuracao" runat="server" Width="618px">
                        <asp:ListItem></asp:ListItem>
                        <asp:ListItem>A - Anual</asp:ListItem>
                        <asp:ListItem>T - Trimestral</asp:ListItem>
                    </asp:DropDownList>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Forma de Tributação:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlFormaDeTributacao" runat="server" Width="618px">
                        <asp:ListItem></asp:ListItem>
                        <asp:ListItem>1 - Real</asp:ListItem>
                        <asp:ListItem>2 - Real Arbitrado</asp:ListItem>
                        <asp:ListItem>3 - Real Presumido (Trimestral)</asp:ListItem>
                        <asp:ListItem>4 - Real Presumido Arbitrado (Trimestral)</asp:ListItem>
                    </asp:DropDownList>
                </div>
            </div>
            <div class="row" style="font-weight: bold;">
                <div class="coltxt" style="margin-left: 178px; width: 152px;">
                    1º Trimestre
                </div>
                <div class="coltxt" style="width: 152px;">
                    2º Trimestre
                </div>
                <div class="coltxt" style="width: 152px;">
                    3º Trimestre
                </div>
                <div class="coltxt" style="width: 152px;">
                    4º Trimestre
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Trimestre de Lucro Arbitrado:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlTrimestre_01" runat="server" Width="152px">
                        <asp:ListItem></asp:ListItem>
                        <asp:ListItem>0 - N&#227;o Arbitrado</asp:ListItem>
                        <asp:ListItem>1 - Arbitrado</asp:ListItem>
                    </asp:DropDownList>
                    <asp:DropDownList ID="ddlTrimestre_02" runat="server" Width="152px">
                        <asp:ListItem></asp:ListItem>
                        <asp:ListItem>0 - N&#227;o Arbitrado</asp:ListItem>
                        <asp:ListItem>1 - Arbitrado</asp:ListItem>
                    </asp:DropDownList>
                    <asp:DropDownList ID="ddlTrimestre_03" runat="server" Width="152px">
                        <asp:ListItem></asp:ListItem>
                        <asp:ListItem>0 - N&#227;o Arbitrado</asp:ListItem>
                        <asp:ListItem>1 - Arbitrado</asp:ListItem>
                    </asp:DropDownList>
                    <asp:DropDownList ID="ddlTrimestre_04" runat="server" Width="152px">
                        <asp:ListItem></asp:ListItem>
                        <asp:ListItem>0 - N&#227;o Arbitrado</asp:ListItem>
                        <asp:ListItem>1 - Arbitrado</asp:ListItem>
                    </asp:DropDownList>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Apuração do Trimestre:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlApuracaoDoTrimestre_01" runat="server" Width="152px">
                        <asp:ListItem></asp:ListItem>
                        <asp:ListItem>0 - Fora P. Esc.</asp:ListItem>
                        <asp:ListItem>1 - Real</asp:ListItem>
                        <asp:ListItem>2 - Arbitrado</asp:ListItem>
                        <asp:ListItem>3 - Presumido</asp:ListItem>
                        <asp:ListItem>4 - Inativo</asp:ListItem>
                    </asp:DropDownList>
                    <asp:DropDownList ID="ddlApuracaoDoTrimestre_02" runat="server" Width="152px">
                        <asp:ListItem></asp:ListItem>
                        <asp:ListItem>0 - Fora P. Esc.</asp:ListItem>
                        <asp:ListItem Value="1 - Real"></asp:ListItem>
                        <asp:ListItem>2 - Arbitrado</asp:ListItem>
                        <asp:ListItem>3 - Presumido</asp:ListItem>
                        <asp:ListItem>4 - Inativo</asp:ListItem>
                    </asp:DropDownList>
                    <asp:DropDownList ID="ddlApuracaoDoTrimestre_03" runat="server" Width="152px">
                        <asp:ListItem></asp:ListItem>
                        <asp:ListItem>0 - Fora P. Esc.</asp:ListItem>
                        <asp:ListItem>1 - Real</asp:ListItem>
                        <asp:ListItem>2 - Arbitrado</asp:ListItem>
                        <asp:ListItem>3 - Presumido</asp:ListItem>
                        <asp:ListItem>4 - Inativo</asp:ListItem>
                    </asp:DropDownList>
                    <asp:DropDownList ID="ddlApuracaoDoTrimestre_04" runat="server" Width="152px">
                        <asp:ListItem></asp:ListItem>
                        <asp:ListItem>0 - Fora P. Esc.</asp:ListItem>
                        <asp:ListItem>1 - Real</asp:ListItem>
                        <asp:ListItem>2 - Arbitrado</asp:ListItem>
                        <asp:ListItem>3 - Presumido</asp:ListItem>
                        <asp:ListItem>4 - Inativo</asp:ListItem>
                    </asp:DropDownList>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Arquivo de Saída:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtArquivodeSaida" runat="server" Width="608px" data-ToolTip="default"
                        ToolTip="É o local/nome do arquivo texto gerado pelo processo que será importado pelo validador ECD Contábil." />
                </div>
                <div class="coltxt" style="float: right;">
                    <asp:UpdatePanel ID="updpnlArquivoDeSaida" runat="server">
                        <ContentTemplate>
                            <asp:Button ID="cmdArquivoDeSaida" CssClass="botao" runat="server" Height="20px"
                                Text="Download" Visible="false" />
                        </ContentTemplate>
                        <Triggers>
                            <asp:PostBackTrigger ControlID="cmdArquivoDeSaida" />
                        </Triggers>
                    </asp:UpdatePanel>
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
