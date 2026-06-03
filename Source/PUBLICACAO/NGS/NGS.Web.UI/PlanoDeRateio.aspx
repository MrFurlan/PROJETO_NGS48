<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="PlanoDeRateio.aspx.vb" Inherits="NGS.Web.UI.PlanoDeRateio" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngPlanoDeRateio" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlPlanoDeRateio" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Plano de Rateio
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
                        <li class="iconConsultar" runat="server">
                            <asp:LinkButton ID="lnkConsultar" Text="Consultar" runat="server" />
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
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoEmpresa" runat="server" />
                    <asp:TextBox ID="txtEmpresa" runat="server" Width="540px" Enabled="False" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnEmpresa" runat="server" Text=">" CssClass="btn" UseSubmitBehavior="False"
                        Enabled="False" OnClick="btnEmpresa_Click" data-ToolTip="default" ToolTip="Empresa Para negociação/consulta." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Grupo:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlGrupoProduto" runat="server" Width="580px" Enabled="False"
                        AutoPostBack="True" OnSelectedIndexChanged="ddlGrupoProduto_SelectedIndexChanged" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Produto:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlProduto" runat="server" Width="580px" Enabled="False" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Percentual:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPercentual" runat="server" Enabled="False" data-ToolTip="default"
                        ToolTip="Porcentagem de rateio." />
                </div>
                <div class="coltxt">
                    <asp:ImageButton ID="imgIncluir" runat="server" Enabled="False" ImageAlign="AbsMiddle"
                        ImageUrl="~/images/ico-mais.gif" OnClick="imgIncluir_Click" data-ToolTip="default"
                        ToolTip="Adicionar ítem" CssClass="btn" Width="18px" />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="gridPlanoDeRateio" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%" OnRowCommand="gridPlanoDeRateio_RowCommand">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:TemplateField>
                            <ItemTemplate>
                                <asp:ImageButton ID="imgExcluir" runat="server" CommandArgument='<%# Eval("Indice") %>'
                                    CommandName="Select" Height="18px" ImageUrl="~/images/erro.jpg" Width="18px" />
                            </ItemTemplate>
                            <HeaderStyle Width="30px" />
                            <ItemStyle Width="30px" />
                        </asp:TemplateField>
                        <asp:BoundField HeaderText="Empresa" DataField="Empresa">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField HeaderText="Grupo" DataField="Grupo">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField HeaderText="Produto" DataField="Produto">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField HeaderText="Percentual" DataField="Percentual" DataFormatString="{0:N2}">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaEmpresas ID="ucConsultaEmpresas" runat="server" />
</asp:Content>
