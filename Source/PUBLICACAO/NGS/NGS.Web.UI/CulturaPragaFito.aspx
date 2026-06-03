<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="CulturaPragaFito.aspx.vb" Inherits="NGS.Web.UI.CulturaPragaFito" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <script type="text/javascript">
        function ValidarObservacao(event) {
            if (document.all) {
                event.returnValue = false;
            }
        }
    </script>
    <style type="text/css">
        .painelleft {
            width: 49%;
            margin-right: 4px;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngCulturaPragaFito" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlCulturaPragaFito" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                CulturaXPragaXFito
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconNovo" ID="lnkNovo" Text="Gravar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAtualizar" ID="lnkAtualizar" Text="Atualizar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconExcluir" ID="lnkExcluir" Text="Excluir" runat="server"
                                OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconConsultar" ID="lnkConsultar" Text="Consultar" runat="server" />
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
            <div class="row" style="line-height: 14px;">
                <ajaxToolkit:TabContainer ID="TabContainer1" runat="server" ActiveTabIndex="0">
                    <ajaxToolkit:TabPanel ID="TabPanel1" runat="server" HeaderText="Cadastro">
                        <ContentTemplate>
                            <div class="row">
                                <div class="collbl">
                                    Código:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtCulturaPragaFito" runat="server" Enabled="False" data-ToolTip="default" ToolTip="Número de cadastro da cultura." />
                                </div>
                                <div class="collbl">
                                    Situação:
                                </div>
                                <div class="coltxt">
                                    <asp:RadioButton ID="radAtivo" runat="server" Enabled="False" GroupName="status"
                                        Text="Ativo" Font-Size="9pt" data-ToolTip="default" ToolTip="Indica se o ítem está ativo ou inativo." />
                                </div>
                                <div class="coltxt">
                                    <asp:RadioButton ID="radBaixado" runat="server" Enabled="False" GroupName="status"
                                        Text="Inativo" Font-Size="9pt" data-ToolTip="default" ToolTip="Indica se o ítem está ativo ou inativo." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Cultura:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlCultura" runat="server" Width="500px" Enabled="False" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Praga:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlPraga" runat="server" Width="500px" Enabled="False" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Fito:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlFito" runat="server" Width="500px" Enabled="False" OnSelectedIndexChanged="ddlFito_SelectedIndexChanged"
                                        AutoPostBack="True" />
                                </div>
                            </div>
                        </ContentTemplate>
                    </ajaxToolkit:TabPanel>
                    <ajaxToolkit:TabPanel ID="TabPanel3" runat="server" HeaderText="Consulta">
                        <ContentTemplate>
                            <div class="bordagrid">
                                <asp:GridView ID="gridCulturaPragaFito" runat="server" AutoGenerateColumns="False"
                                    CellPadding="4" ForeColor="#333333" GridLines="None" OnSelectedIndexChanged="gridCulturaPragaFito_SelectedIndexChanged"
                                    Width="100%">
                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <EditRowStyle BackColor="#999999" />
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <Columns>
                                        <asp:CommandField ButtonType="Button" SelectText=" &gt;" ShowSelectButton="True">
                                            <HeaderStyle Width="30px" />
                                            <ItemStyle Width="30px" />
                                        </asp:CommandField>
                                        <asp:BoundField DataField="CulturaPragaFito_Id" HeaderText="Código">
                                            <HeaderStyle HorizontalAlign="Left" Width="50px" />
                                            <ItemStyle HorizontalAlign="Left" Width="50px" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Cultura" HeaderText="Cultura">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Praga" HeaderText="Praga">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Fito" HeaderText="Fito">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                    </Columns>
                                </asp:GridView>
                            </div>
                        </ContentTemplate>
                    </ajaxToolkit:TabPanel>
                </ajaxToolkit:TabContainer>
                <div class="subtitulodiv">
                    Dosagem
                </div>
                <div class="row">
                    <div class="collbl">
                        Solo:
                    </div>
                    <div class="coltxt" style="width: 303px;">
                        <asp:DropDownList ID="ddlSolo" runat="server" Width="260px" Enabled="False" />
                    </div>
                    <div class="collbl" style="width: 150px;">
                        Recomendada:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtRecomendada" runat="server" CssClass="txtDecimal4" Width="50px"
                            Enabled="False" data-ToolTip="default" ToolTip="" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Mínima:
                    </div>
                    <div class="coltxt" style="width: 303px;">
                        <asp:TextBox ID="txtMinimo" runat="server" Width="50px" Enabled="False" CssClass="txtDecimal4"
                            data-ToolTip="default" ToolTip="" />
                    </div>
                    <div class="collbl" style="width: 150px;">
                        Máxima:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtMaximo" runat="server" Width="50px" Enabled="False" CssClass="txtDecimal4"
                            data-ToolTip="default" ToolTip="" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Vazão terrestre:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtVazaoTerrestre" runat="server" Width="293px" Enabled="False"
                            MaxLength="100" data-ToolTip="default" ToolTip="" />
                    </div>
                    <div class="collbl" style="width: 150px;">
                        Vazão aérea:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtVazaoAerea" runat="server" Width="99%" Enabled="False" MaxLength="100"
                            data-ToolTip="default" ToolTip="" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Unidade de Medida:
                    </div>
                    <div class="coltxt" style="width: 303px;">
                        <asp:TextBox ID="txtUnidadeDeMedida" runat="server" data-ToolTip="default" ToolTip="" />
                    </div>
                    <div class="collbl" style="width: 150px;">
                        Intervalo de Segurança:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtIntervaloDeSeguranca" runat="server" Enabled="False" MaxLength="50"
                            data-ToolTip="default" ToolTip="" />
                    </div>
                    <div class="coltxt">
                        <asp:Button ID="btnAdicionarDosagem" runat="server" CssClass="botao" Style="width: 130px !important; height: 25px !important; font-size: 8pt !important"
                            Text="Adicionar Dosagem" data-ToolTip="default" ToolTip="" />
                    </div>
                </div>
                <div class="bordagrid">
                    <asp:GridView ID="gridDosagem" runat="server" Width="98%" OnRowCommand="gridDosagem_RowCommand"
                        AutoGenerateColumns="False" GridLines="None" ForeColor="#333333" CellPadding="4">
                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                        <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <EditRowStyle BackColor="#999999" />
                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                        <Columns>
                            <asp:TemplateField>
                                <ItemTemplate>
                                    <asp:ImageButton ID="imgExcluir" runat="server" Height="18px" ImageUrl="~/images/erro.jpg"
                                        Width="18px" CommandArgument='<%# Eval("Indice") %>' CommandName="Select" />
                                </ItemTemplate>
                                <HeaderStyle Width="30px" />
                                <ItemStyle Width="30px" />
                            </asp:TemplateField>
                            <asp:BoundField DataField="Solo" HeaderText="Solo" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Minima" HeaderText="M&#237;nima" DataFormatString="{0:N4}"
                                HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Recomendada" HeaderText="Recomendada" DataFormatString="{0:N4}"
                                HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Maxima" HeaderText="M&#225;xima" DataFormatString="{0:N4}"
                                HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:BoundField>
                            <asp:BoundField DataField="UnidadeDeMedida" HeaderText="Und.Medida" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="VazaoTerrestre" HeaderText="Vaz&#227;o Terrestre" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="VazaoAerea" HeaderText="Vaz&#227;o A&#233;rea" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="IntervaloDeSeguranca" HeaderText="Intervalo de Seguran&#231;a"
                                HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:BoundField>
                        </Columns>
                    </asp:GridView>
                </div>
                <div class="row">
                    <div class="subtitulodiv">
                        Aplicação
                    </div>
                </div>
                <div class="painelleft">
                    <div class="collbl">
                        Época:
                    </div>
                    </br>
                    <div class="coltxt" style="margin-top: 4px;">
                        <asp:TextBox ID="txtEpocaDeAplicacao" runat="server" Enabled="False" Height="310px"
                            TextMode="MultiLine" Width="445px" data-ToolTip="default" ToolTip="Detalhamento da aplicação do defensivo." />
                    </div>
                </div>
                <div class="painelleft">
                    <div class="collbl">
                        Forma:
                    </div>
                    <div class="bordagrid" style="margin-top: 32px; height: 310px">
                        <asp:GridView ID="gridFormaDeAplicacao" runat="server" AutoGenerateColumns="False"
                            overflow="auto" CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                            <EditRowStyle BackColor="#999999" />
                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                            <Columns>
                                <asp:TemplateField>
                                    <ItemTemplate>
                                        <asp:CheckBox ID="chkCodigo" runat="server" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:BoundField DataField="Codigo" HeaderText="Código" ControlStyle-CssClass="none"
                                    HeaderStyle-CssClass="none" ItemStyle-CssClass="none">
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemStyle HorizontalAlign="Left" />
                                </asp:BoundField>
                                <asp:BoundField DataField="Descricao" HeaderText="Descrição">
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemStyle HorizontalAlign="Left" />
                                </asp:BoundField>
                                <asp:TemplateField HeaderText="Modo de Aplicação">
                                    <ItemTemplate>
                                        <asp:TextBox ID="txtModoDeAplicacao" runat="server" Height="15px" Width="98%" />
                                    </ItemTemplate>
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemStyle HorizontalAlign="Left" Width="70%" />
                                </asp:TemplateField>
                            </Columns>
                        </asp:GridView>
                    </div>
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
