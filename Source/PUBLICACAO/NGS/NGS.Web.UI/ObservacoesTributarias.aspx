<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="ObservacoesTributarias.aspx.vb" Inherits="NGS.Web.UI.ObservacoesTributarias" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:ScriptManager ID="scrmngObservacoesTributarias" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" />
    <asp:UpdatePanel ID="updpnlObservacoesTributarias" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Observações Tributárias
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
                    <asp:TextBox ID="txtCodigo" runat="server" Width="170px" Enabled="False" data-ToolTip="default"
                        ToolTip="Código de cadastro da observação tributária." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Estado:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEstado" runat="server" AutoPostBack="True" Width="300px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Encargo:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEncargos" runat="server" Width="600px" AutoPostBack="true" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Descrição:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtdescricao" runat="server" Width="591px" data-ToolTip="default"
                        ToolTip="Detalhamento do encargo." TextMode="MultiLine" />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="gridObsTrib" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="gridObsTrib_SelectedIndexChanged">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                        <asp:BoundField DataField="Codigo" HeaderText="Código">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Descricao" HeaderText="Descrição">
                            <HeaderStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:TemplateField HeaderText="Estado" HeaderStyle-HorizontalAlign="Left" ItemStyle-Width="100px">
                            <ItemTemplate>
                                <%# Eval("Estado.Descricao") %>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Encargo" HeaderStyle-HorizontalAlign="Left" ItemStyle-Width="100px">
                            <ItemTemplate>
                                <%# Eval("CodigoEncargo")%>
                            </ItemTemplate>
                        </asp:TemplateField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
