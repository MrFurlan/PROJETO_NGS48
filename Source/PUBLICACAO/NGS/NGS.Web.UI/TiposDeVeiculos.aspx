<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="TiposDeVeiculos.aspx.vb" Inherits="NGS.Web.UI.TiposDeVeiculos" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngTiposDeVeiculos" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlTiposDeVeiculos" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Tipos de Veículos
            </div>
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
                    <asp:TextBox ID="txtCodigo" TabIndex="1" runat="server" Width="100px" Enabled="False"
                        CssClass="txtNumerico" data-ToolTip="default" ToolTip="Código de cadastro dos tipos de veículos." />
                </div>
                <div class="collbl">
                    Descrição:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDescricao" TabIndex="2" runat="server" Width="500px" MaxLength="100"
                        data-ToolTip="default" ToolTip="Descrição do veículo." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Capacidade:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCapacidade" TabIndex="3" runat="server" Width="100px" data-ToolTip="default"
                        ToolTip="Capacidade de peso total do veículo." />
                </div>
                <div class="collbl">
                    Via de Transporte:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlViaTransporte" runat="server" TabIndex="4" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Tara Mínima:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtTaraMinima" TabIndex="5" runat="server" Width="100px" CssClass="txtNumerico"
                        data-ToolTip="default" ToolTip="Peso mínimo que o veículo pode carregar." />
                </div>
                <div class="collbl">
                    Tara Máxima:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtTaraMaxima" TabIndex="6" runat="server" Width="100px" CssClass="txtNumerico"
                        data-ToolTip="default" ToolTip="Peso máximo que o veículo pode carregar." />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="GridTiposDeVeiculos" runat="server" AutoGenerateColumns="False"
                    CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="GridTiposDeVeiculos_SelectedIndexChanged">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:CommandField SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left" Width="25px"></ItemStyle>
                        </asp:CommandField>
                        <asp:BoundField DataField="Codigo" HeaderText="C&#243;digo">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left" Width="60px"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Descricao" HeaderText="Descri&#231;&#227;o">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Capacidade" DataFormatString="{0:N0}" HeaderText="Capacidade">
                            <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Right" Width="80px"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="TaraMinima" DataFormatString="{0:N0}" HeaderText="Tara M&#237;nima">
                            <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Right" Width="80px"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="TaraMaxima" DataFormatString="{0:N0}" HeaderText="Tara M&#225;xima">
                            <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Right" Width="80px"></ItemStyle>
                        </asp:BoundField>
                        <asp:TemplateField>
                            <ItemTemplate>
                                <asp:HiddenField ID="txtCodigoVia" runat="server" Value='<%# Eval("ViaDeTransporte") %>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Via de Transporte" HeaderStyle-HorizontalAlign="Left"
                            ItemStyle-Width="100px">
                            <ItemTemplate>
                                <%# IIf(String.IsNullOrEmpty(Eval("ViaDeTransporteDetalhes.Descricao")), "", Eval("ViaDeTransporteDetalhes.Descricao"))%>
                            </ItemTemplate>
                        </asp:TemplateField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
