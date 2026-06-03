<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="SpedFiscal.aspx.vb" Inherits="NGS.Web.UI.SpedFiscal" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <script type="text/javascript">
        function downloadArquivo() {
            alert("Processo Concluido");
            //msgbox("Processo Concluido", "SUCESSO!", "Sucess");
            $("#MainContent_imdDownload").click();
        };
    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngSpedFiscal" runat="server" AsyncPostBackTimeout="50000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updnlSpedFiscal" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Sped Fiscal
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconRelatorio" runat="server">
                            <asp:LinkButton ID="lnkNovo" Text="Processar" runat="server" />
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
                    <asp:DropDownList ID="ddlUnidade" runat="server" Width="634px" OnSelectedIndexChanged="DdlUnidade_SelectedIndexChanged"
                        AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlEmpresa" runat="server" Width="634px" OnSelectedIndexChanged="DdlEmpresa_SelectedIndexChanged"
                        AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    ICMS:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlProcesso" runat="server" Width="634px" OnSelectedIndexChanged="DdlProcesso_SelectedIndexChanged"
                        AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    IPI:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlProcessoIPI" runat="server" Width="634px" OnSelectedIndexChanged="DdlProcesso_SelectedIndexChanged"
                        AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Período:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicial" CssClass="calendario" runat="server" Width="75px" data-ToolTip="default" ToolTip="Informar o período de apuração." />
                    &nbsp;à&nbsp;
                    <asp:TextBox ID="txtDataFinal" runat="server" CssClass="calendario" Width="75px" data-ToolTip="default" ToolTip="Informar o período de apuração." />
                </div>
                <div class="collbl">
                    Livro:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtLivro" runat="server" Width="64px" CssClass="txtNumerico" data-ToolTip="default" ToolTip="Número do livro." />
                </div>
                <div class="collbl">
                    Folha:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtFolha" runat="server" Width="64px" CssClass="txtNumerico" data-ToolTip="default" ToolTip="Número do folha." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Finalidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlFinalidade" runat="server" Width="634px">
                        <asp:ListItem Value="0">0 - Remessa do Arquivo Original</asp:ListItem>
                        <asp:ListItem Value="1">1 - Remessa do Arquivo Substituto</asp:ListItem>
                    </asp:DropDownList>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Perfil:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlPerfil" runat="server" Width="634px">
                        <asp:ListItem Value="A">A - Perfil A</asp:ListItem>
                        <asp:ListItem Value="B">B - Perfil B</asp:ListItem>
                        <asp:ListItem Value="C">C - Perfil C</asp:ListItem>
                    </asp:DropDownList>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Atividade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlAtividade" runat="server" Width="634px">
                        <asp:ListItem Value="0">0 - Industrial ou equiparado a industria</asp:ListItem>
                        <asp:ListItem Value="1">1 - Outros</asp:ListItem>
                    </asp:DropDownList>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Arq. Saída:
                </div>
                <div class="coltxt">
                    <asp:UpdatePanel ID="updpnlArquivoDeSaida" runat="server">
                        <ContentTemplate>
                            <asp:TextBox ID="txtArquivoDeSaida" runat="server" Width="630px" Enabled="False" data-ToolTip="default" ToolTip="Nome do arquivo do SPED." />
                            <asp:ImageButton ID="imdDownload" runat="server" ImageAlign="AbsMiddle" ImageUrl="Images/download32x32.png"
                                Style="margin-top: 0;" OnClick="imdDownload_Click" data-ToolTip="default" ToolTip="Baixar Arquivo"
                                Height="22px" Width="22px" />
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
