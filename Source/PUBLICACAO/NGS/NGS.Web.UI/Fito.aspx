<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="Fito.aspx.vb" Inherits="NGS.Web.UI.Fito" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngFito" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlFito" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Fitossanitário
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
            <div class="row">
                <div class="collbl">
                    Código:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCodigo" runat="server" Enabled="False" />
                </div>
                <div class="collbl">
                    INDEA:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCodigoIndea" runat="server" MaxLength="50" Enabled="False" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Classe Toxicolôgica:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlClasseToxicologica" runat="server" Width="558px" Enabled="False" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Classe Ambiental:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlClasseAmbiental" runat="server" Width="558px" Enabled="False" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Formulação do Fito:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlFormulacaoDoFito" runat="server" Width="558px" Enabled="False" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Ingrediente Ativo:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlIngredienteAtivo" runat="server" Width="558px" Enabled="False" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Classe de Risco:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlClasseDeRisco" runat="server" Width="558px" Enabled="False" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Nome Comercial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNomeComercial" runat="server" MaxLength="50" Width="547px" Enabled="False" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Nome Técnico:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNomeTecnico" runat="server" MaxLength="50" Width="547px" Enabled="False" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Formula Bruta:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtFormulaBruta" runat="server" MaxLength="50" Width="547px" Enabled="False" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Concentração:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtConcentracao" runat="server" MaxLength="50" Width="547px" Enabled="False" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    UF Restrição:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtUFRestricao" runat="server" MaxLength="20" Enabled="False" />
                </div>
                <div class="collbl">
                    Inflamável:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="radInflamavelSim" runat="server" GroupName="Inflamavel" Text="Sim"
                        ValidationGroup="Inflamavel" Enabled="False" />
                    <asp:RadioButton ID="radInflamavelNao" runat="server" GroupName="Inflamavel" Text="Não"
                        ValidationGroup="Inflamavel" Enabled="False" />
                </div>
                <div class="collbl">
                    Corrosivo:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="radCorrosivoSim" runat="server" GroupName="Corrosivo" Text="Sim"
                        ValidationGroup="Corrosivo" Enabled="False" />
                    <asp:RadioButton ID="radCorrosivoNao" runat="server" GroupName="Corrosivo" Text="Não"
                        ValidationGroup="Corrosivo" Enabled="False" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    MA:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtMA" runat="server" MaxLength="50" Width="547px" Enabled="False" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    ONU:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtOnu" runat="server" MaxLength="50" Width="547px" Enabled="False" />
                </div>
            </div>
            <ajaxToolkit:TabContainer ID="TabContainer1" runat="server" ActiveTabIndex="1" Width="100%">
                <ajaxToolkit:TabPanel runat="server" ID="TabPanel1">
                    <HeaderTemplate>
                        Fitossanitários
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="bordagrid" style="height: 180px;">
                            <asp:Panel ID="pnlLista" runat="server" Height="180px" ScrollBars="Vertical" Width="100%">
                                <asp:GridView ID="gridFitossanitario" runat="server" CellPadding="4" ForeColor="#333333"
                                    GridLines="None" Width="98%" AutoGenerateColumns="False" OnSelectedIndexChanged="gridFitossanitario_SelectedIndexChanged">
                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <EditRowStyle BackColor="#999999" />
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <Columns>
                                        <asp:CommandField ButtonType="Button" SelectText=" &gt;" ShowSelectButton="True" />
                                        <asp:BoundField DataField="CodigoFito" HeaderText="Fito">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="CodigoIndeaMT" HeaderText="Indea">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="CodigoClasseTox" HeaderText="C.Tox.">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="CodigoClasseAmbiental" HeaderText="C.Amb.">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="CodigoIA" HeaderText="IA">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="CodigoClasseDeRisco" HeaderText="C.Risco">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="RegistroMA" HeaderText="Registro MA">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="RegistroONU" HeaderText="Registro ONU">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                    </Columns>
                                </asp:GridView>
                            </asp:Panel>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" ID="TabPanel2">
                    <HeaderTemplate>
                        Modo de Aplicação
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="row">
                            <div class="coltxt lg">
                                <asp:TextBox ID="txtModoDeAplicacao" runat="server" MaxLength="547" TextMode="MultiLine"
                                    Rows="10" Enabled="False" Style="width: 99%;" />
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" ID="TabPanel3">
                    <HeaderTemplate>
                        Instruções de Uso
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="row">
                            <div class="coltxt lg">
                                <asp:TextBox ID="txtUso" runat="server" Height="180px" MaxLength="547" TextMode="MultiLine"
                                    Rows="10" Width="100%" Enabled="False" />
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" ID="TabPanel4">
                    <HeaderTemplate>
                        Descarte da Embalagem
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="row">
                            <div class="coltxt lg">
                                <asp:TextBox ID="txtDescarte" runat="server" Height="180px" MaxLength="547" TextMode="MultiLine"
                                    Rows="10" Width="99%" Enabled="False" />
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" ID="TabPanel5">
                    <HeaderTemplate>
                        Primeiros Socorros
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="row">
                            <div class="coltxt lg">
                                <asp:TextBox ID="txtPrimeirosSocorros" runat="server" Height="180px" MaxLength="547"
                                    Rows="10" TextMode="MultiLine" Width="99%" Enabled="False" />
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" ID="TabPanel6">
                    <HeaderTemplate>
                        Meio Ambiente
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="row">
                            <div class="coltxt lg">
                                <asp:TextBox ID="txtMeioAmbiente" runat="server" Height="180px" MaxLength="547" TextMode="MultiLine"
                                    Width="99%" Enabled="False" />
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" ID="TabPanel7">
                    <HeaderTemplate>
                        Incompatibilidade
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="row">
                            <div class="coltxt lg">
                                <asp:TextBox ID="txtIncompatibilidade" runat="server" Height="180px" MaxLength="547"
                                    TextMode="MultiLine" Width="99%" Enabled="False" />
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
