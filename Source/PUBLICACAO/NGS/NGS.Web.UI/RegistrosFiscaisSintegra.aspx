<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="RegistrosFiscaisSintegra.aspx.vb" Inherits="NGS.Web.UI.RegistrosFiscaisSintegra" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        .collbl {
            width: 140px;
        }
    </style>
    <script type="text/javascript">
        function downloadArquivo() {
            $("img.download").each(function () {
                $(this).click();
            });
        }
    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngRegistrosFiscaisSintegra" EnableScriptGlobalization="true"
        EnableScriptLocalization="true" runat="server" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlRegistrosFiscaisSintegra" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Registros Fiscais Sintegra
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconNovo" runat="server">
                            <asp:UpdatePanel ID="updpnlNovo" runat="server">
                                <ContentTemplate>
                                    <asp:LinkButton ID="lnkNovo" Text="Processar" runat="server" />
                                </ContentTemplate>
                                <Triggers>
                                    <asp:PostBackTrigger ControlID="lnkNovo" />
                                </Triggers>
                            </asp:UpdatePanel>
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" runat="server" Text="Limpar" /></li>
                        <li class="iconAjuda" runat="server">
                            <asp:LinkButton ID="lnkAjuda" Text="Ajuda" runat="server" /></li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUnidade" runat="server" Width="610px" OnSelectedIndexChanged="DdlUnidade_SelectedIndexChanged"
                        AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="610px" AutoPostBack="True"
                        OnSelectedIndexChanged="ddlEmpresa_SelectedIndexChanged" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Processo:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtProcesso" runat="server" Width="88px" CssClass="txtNumerico"
                        OnTextChanged="txtProcesso_TextChanged" MaxLength="8" data-ToolTip="default" ToolTip="Número do processo." />
                </div>
                <div class="collbl">
                    Período:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPeriodoInicial" runat="server" Width="88px" CssClass="calendario" data-ToolTip="default" ToolTip="Informar de acordo com o período de apuração." />
                    <asp:TextBox ID="txtPeriodoFinal" CssClass="calendario" runat="server" Width="88px" data-ToolTip="default" ToolTip="Informar de acordo com o período de apuração." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Convênio:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlConvenio" runat="server" Width="610px">
                        <asp:ListItem Value="1">1 - Estrutura cfe. Conv&#234;nio ICMS 57/95, na vers&#231;&#227;o estabelecida pelo conv&#234;nio ICMS 31/99 e com as altera&#231;&#245;es promovidas até o Conv&#234;nio ICMS 30/02</asp:ListItem>
                        <asp:ListItem Value="2">2 - Estrutura cfe. Conv&#234;nio ICMS 57/95, na vers&#231;&#227;o estabelecida pelo conv&#234;nio ICMS 69/02 e com as altera&#231;&#245;es promovidas pelo Conv&#234;nio ICMS 142/02</asp:ListItem>
                        <asp:ListItem Value="3">3 - Estrutura cfe. Conv&#234;nio ICMS 57/95, com as altera&#231;&#245;es promovidas pelo conv&#234;nio ICMS 76/03</asp:ListItem>
                    </asp:DropDownList>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Natureza da Operação:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlNaturezaDaOperacao" runat="server" Width="610px">
                        <asp:ListItem Value="1">1 - Interestaduais somente opera&#231;&#245;es sujeitas ao regime de Substitui&#231;&#227;o Tribut&#225;ria</asp:ListItem>
                        <asp:ListItem Value="2">2 - Interestaduais - opera&#231;&#245;es com ou sem Substitui&#231;&#227;o Tribut&#225;ria</asp:ListItem>
                        <asp:ListItem Selected="True" Value="3">3 - Totalidade das opera&#231;&#245;es do informante</asp:ListItem>
                    </asp:DropDownList>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Arquivo Magnético:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlArquivoMagnetico" runat="server" Width="610px">
                        <asp:ListItem Value="1">1 - Normal</asp:ListItem>
                        <asp:ListItem Value="2">2 - Retifica&#231;&#227;o Total de Arquivo</asp:ListItem>
                        <asp:ListItem Value="3">3 - Retifica&#231;&#227;o Aditiva do Arquivo</asp:ListItem>
                        <asp:ListItem Value="4">4 - Retifica&#231;&#227;o Aditiva do Arquivo</asp:ListItem>
                        <asp:ListItem Value="5">5 - Desfazimento</asp:ListItem>
                    </asp:DropDownList>
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="grdProcesso" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="grdProcesso_SelectedIndexChanged"
                    DataKeyNames="ArquivoMagnetico">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:CommandField SelectImageUrl="~/Images/seta1.gif" SelectText=" &gt; " ShowSelectButton="True"
                            ButtonType="Button"></asp:CommandField>
                        <asp:BoundField DataField="Processo_Id" HeaderText="Processo"></asp:BoundField>
                        <asp:BoundField DataField="PeriodoInicial" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Per&#237;odo Inicial"
                            HtmlEncode="False"></asp:BoundField>
                        <asp:BoundField DataField="PeriodoFinal" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Per&#237;odo Final"
                            HtmlEncode="False"></asp:BoundField>
                        <asp:BoundField DataField="Convenio" HeaderText="Conv&#234;nio"></asp:BoundField>
                        <asp:BoundField DataField="NaturezaDaOperacao" HeaderText="Natureza"></asp:BoundField>
                        <asp:BoundField ConvertEmptyStringToNull="False" DataField="ArquivoMagnetico" Visible="False">
                            <ControlStyle BorderWidth="0px" Height="0px" Width="0px"></ControlStyle>
                            <FooterStyle BorderWidth="0px" Height="0px" Width="0px"></FooterStyle>
                            <HeaderStyle BorderWidth="0px" Height="0px" Width="0px"></HeaderStyle>
                            <ItemStyle BorderWidth="0px" Height="0px" Width="0px"></ItemStyle>
                        </asp:BoundField>
                        <asp:TemplateField HeaderText="Baixar" HeaderStyle-HorizontalAlign="Center">
                            <ItemTemplate>
                                <asp:UpdatePanel ID="updpnlArquivoDeSaida" runat="server">
                                    <ContentTemplate>
                                        <asp:LinkButton ID="imdBaixarArquivo" runat="server" OnClick="ImdBaixarArquivo_Click">
                                            <asp:Image ID="imgPdf" runat="server" CssClass="download" Width="16px" Height="16px"
                                                ImageUrl="~/Images/bottom.png" data-ToolTip="default" ToolTip="Baixar Arquivo" />
                                        </asp:LinkButton>
                                    </ContentTemplate>
                                    <Triggers>
                                        <asp:PostBackTrigger ControlID="imdBaixarArquivo" />
                                    </Triggers>
                                </asp:UpdatePanel>
                            </ItemTemplate>
                            <ItemStyle Width="70px" HorizontalAlign="Center" VerticalAlign="Middle" />
                        </asp:TemplateField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
