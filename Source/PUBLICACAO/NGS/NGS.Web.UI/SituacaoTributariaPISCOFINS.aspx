<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="SituacaoTributariaPISCOFINS.aspx.vb" Inherits="NGS.Web.UI.SituacaoTributariaPISCOFINS" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngSituacaoTributacaoPISCOFINS" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlSituacaoTributariaPISCOFINS" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Situação Tributária do PIS e COFINS
            </div>
            <ajaxToolkit:TabContainer ID="tcPISCOFINS" runat="server" ActiveTabIndex="0" Width="100%">
                <ajaxToolkit:TabPanel ID="TabPanel1" runat="server">
                    <HeaderTemplate>
                        Situacão
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li class="iconNovo" runat="server">
                                        <asp:LinkButton ID="lnkNovo" Text="Gravar" runat="server" />
                                    </li>
                                    <li class="iconAtualizar" runat="server">
                                        <asp:LinkButton ID="lnkAtualizar" Text="Atualizar" runat="server" />
                                    </li>
                                    <li class="iconExcluir" runat="server">
                                        <asp:LinkButton ID="lnkExcluir" Text="Excluir" runat="server" OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" />
                                    </li>
                                    <li class="iconLimpar" runat="server">
                                        <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                                    </li>
                                    <li class="iconRelatorio" runat="server">
                                        <asp:LinkButton ID="lnkRelatorio" Text="Relatório" runat="server" />
                                    </li>
                                    <li class="iconRelatorio" runat="server" style="width: 18%;">
                                        <asp:LinkButton ID="lnkRelatorioObs" Text="Relatório com Obs." runat="server" />
                                    </li>
                                    <li class="iconAjuda" runat="server">
                                        <asp:LinkButton ID="lnkAjuda" Text="Ajuda" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Código:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtCodigo" runat="server" MaxLength="2" Width="100px" data-ToolTip="default"
                                    ToolTip="Código de cadastro da situação tributária PIS COFINS." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Descrição:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDescricao" runat="server" MaxLength="130" Width="700px" data-ToolTip="default"
                                    ToolTip="Descrição da situação tributária PIS COFINS." />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="GridSTPISCOFINS" runat="server" AutoGenerateColumns="False" CellPadding="3"
                                ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="GridSTPISCOFINS_SelectedIndexChanged">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True">
                                        <ItemStyle Width="25px"></ItemStyle>
                                    </asp:CommandField>
                                    <asp:BoundField DataField="Codigo" HeaderText="Código">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left" Width="50px"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Descricao" HeaderText="Descrição">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel2" runat="server">
                    <HeaderTemplate>
                        Observacões
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li class="iconNovo" runat="server">
                                        <asp:LinkButton ID="lnkNovoObs" Text="Gravar" runat="server" />
                                    </li>
                                    <li class="iconAtualizar" runat="server">
                                        <asp:LinkButton ID="lnkAtualizarObs" Text="Atualizar" runat="server" />
                                    </li>
                                    <li class="iconExcluir" runat="server">
                                        <asp:LinkButton ID="lnkExcluirObs" Text="Excluir" runat="server" OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" />
                                    </li>
                                    <li class="iconLimpar" runat="server">
                                        <asp:LinkButton ID="lnkLimparObs" runat="server" Text="Limpar" />
                                    </li>
                                    <li class="iconAjuda" runat="server">
                                        <asp:LinkButton ID="lnkAjudaObs" runat="server" Text="Ajuda" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                ST PIS COFINS:
                            </div>
                            <div class="coltxt">
                                <asp:Label ID="lblPISCOFINS" runat="server" Width="80%" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Código:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtCodigoObs" runat="server" MaxLength="4" Width="100px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Descrição:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDescricaoObs" runat="server" Width="500px" Height="100px" TextMode="MultiLine" />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="GridSTPISCOFINSObs" runat="server" AutoGenerateColumns="False"
                                CssClass="bordasimples" Width="100%" CellPadding="3" ForeColor="#333333" GridLines="None"
                                OnSelectedIndexChanged="GridSTPISCOFINSObs_SelectedIndexChanged">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True">
                                        <ItemStyle Width="25px"></ItemStyle>
                                    </asp:CommandField>
                                    <asp:BoundField DataField="CodigoObservacao" HeaderText="Código">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left" Width="50px"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Descricao" HeaderText="Descrição">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
