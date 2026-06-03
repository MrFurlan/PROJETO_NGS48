<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="SpedContabil.aspx.vb" Inherits="NGS.Web.UI.SpedContabil" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        .collbl {
            width: 170px;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngSpedContabil" runat="server" AsyncPostBackTimeout="900" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlSpedContabil" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Sped Contábil - ECD
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconRelatorio" runat="server">
                            <asp:LinkButton ID="lnkProcessar" Text="Processar" OnClick="lnkProcessar_Click" runat="server" />
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
                        Width="550px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlEmpresa" runat="server" Width="550px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Periodo:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicial" CssClass="calendario" Width="86px" runat="server"
                        data-ToolTip="default" ToolTip="Data inicial a final para consulta." />
                    &nbsp;a&nbsp;
                    <asp:TextBox ID="txtDataFinal" runat="server" Width="86px" CssClass="calendario"
                        data-ToolTip="default" ToolTip="Data inicial a final para consulta." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Forma de Tributação:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlFormaDeTributacao" runat="server" Width="550px">
                        <asp:ListItem Value="1"> 1 - PJ em Geral – Lucro Real</asp:ListItem>
                        <asp:ListItem Value="2"> 2 - PJ em Geral – Lucro Presumido</asp:ListItem>
                        <asp:ListItem Value="3"> 3 - Financeiras – Lucro Real</asp:ListItem>
                        <asp:ListItem Value="4"> 4 - Seguradoras – Lucro Real</asp:ListItem>
                        <asp:ListItem Value="5"> 5 - Imunes e Isentas em Geral</asp:ListItem>
                        <asp:ListItem Value="6"> 6 - Imunes e Isentas – Financeiras</asp:ListItem>
                        <asp:ListItem Value="7"> 7 - Imunes e Isentas – Seguradoras</asp:ListItem>
                        <asp:ListItem Value="8"> 8 - Entidades Fechadas de Previdência Complementar</asp:ListItem>
                        <asp:ListItem Value="9"> 9 - Partidos Políticos</asp:ListItem>
                        <asp:ListItem Value="10">10 - Financeiras – Lucro Presumido</asp:ListItem>
                    </asp:DropDownList>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Situação Especial:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlSituacaoEspecial" runat="server" Width="550px">
                        <asp:ListItem Value="">Sem Situação Especial</asp:ListItem>
                        <asp:ListItem Value="0">0 - Abertura</asp:ListItem>
                        <asp:ListItem Value="1">1 - Cis&#227;o</asp:ListItem>
                        <asp:ListItem Value="2">2 - Fus&#227;o</asp:ListItem>
                        <asp:ListItem Value="3">3 - Incorpora&#231;&#227;o</asp:ListItem>
                        <asp:ListItem Value="4">4 - Extins&#227;o</asp:ListItem>
                    </asp:DropDownList>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Situação no Início do Período:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlSituacaoInicioPeriodo" runat="server" Width="550px">
                        <asp:ListItem Value="0">0 - Normal (Início no primeiro dia do ano ou do mês)</asp:ListItem>
                        <asp:ListItem Value="1">1 - Abertura</asp:ListItem>
                        <asp:ListItem Value="2">2 - Resultante de cisão/fusão ou remanescente de cisão, ou realizou incorporação</asp:ListItem>
                        <asp:ListItem Value="3">3 - Início de obrigatoriedade da entrega da ECD no curso do ano calendário</asp:ListItem>
                    </asp:DropDownList>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Indicador de Movimento:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlIndicadorDeMovimento" runat="server" Width="550px">
                        <asp:ListItem Value="0">0 - Bloco Com Dados Informados</asp:ListItem>
                        <asp:ListItem Value="1">1 - Bloco sem dados informados</asp:ListItem>
                    </asp:DropDownList>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Natureza do Livro Diário:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlNaturezaDoLivro" runat="server" Width="550px">
                        <asp:ListItem>livro Diario</asp:ListItem>
                    </asp:DropDownList>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Número de Ordem do Livro:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNumeroDoLivro" runat="server" CssClass="txtNumerico" data-ToolTip="default"
                        ToolTip="Número sequencial do livro." />
                </div>
                <div class="collbl">
                    <asp:CheckBox ID="cbTrimestral" runat="server" Text="Apuração do IRPJ Trimestral"
                        data-ToolTip="default" ToolTip="Selecionar se o IRPJ é trimestral." />
                </div>
                <div class="collbl">
                    <asp:CheckBox ID="chkRetifica" runat="server" Text="Retificação"
                        data-ToolTip="default" ToolTip="Retificação" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Arquivo de Saida:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtArquivoDeSaida" Enable="False" runat="server" Width="550px" Enabled="False"
                        data-ToolTip="default" ToolTip="É o local/nome do arquivo texto gerado pelo processo que será importado pelo validador ECD Contábil" />
                </div>
                <div class="coltxt">
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
            <div class="row">
                <div class="collbl">
                    Arquivo Auxiliar:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtArquivoAuxiliar" runat="server" Width="550px" data-ToolTip="default"
                        ToolTip="É o arquivo dos demonstrativos contábeis preparado pelo contador em formato (RTF) que será importado no processo de geração para ser incorporado no arquivo texto gerado pelo processo." />
                </div>
                <div class="coltxt">
                    <asp:UpdatePanel ID="updpnlArquivoDeSaidaAux" runat="server">
                        <ContentTemplate>
                            <asp:Button ID="cmdArquivoAuxDeSaida" runat="server" Height="20px" Text="Download"
                                CssClass="botao" Visible="false" />
                        </ContentTemplate>
                        <Triggers>
                            <asp:PostBackTrigger ControlID="cmdArquivoAuxDeSaida" />
                        </Triggers>
                    </asp:UpdatePanel>
                </div>
            </div>
            <%--<div class="row">
                <div class="collbl">
                    Plano De Contas:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rbPlanoNovo" runat="server" Checked="True" GroupName="plano"
                        Text="Novo" data-ToolTip="default" ToolTip="Selecionar o plano de contas." />
                </div>--%>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
