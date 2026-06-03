<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucFile.ascx.vb" Inherits="NGS.Web.UI.ucFile" %>
<script type="text/javascript">
    function Arquivo() {
        document.getElementById("btnAdicionar_ucFile").click();
        //$("#ctl00$MainContent$ucFile$btnAdicionar").click();
    }
</script>
<asp:UpdatePanel ID="udpFile" runat="server">
    <ContentTemplate>
        <table border="0">
            <tr>
                <td>
                    <asp:FileUpload ID="fupArquivo" onchange="this.form.submit();" runat="server" Width="120px" Font-Size="11px" ClientIDMode="Static" />
                </td>
                <td>
                    <asp:DataList runat="server" ID="dlsArquivo" GridLines="None" RepeatDirection="Horizontal"
                        OnItemDataBound="Item_Bound" OnItemCommand="Item_Command">
                        <ItemTemplate>
                            <div class="none">
                                <asp:Label ID="lblCodigo" runat="server" Text='<%# Eval("Codigo") %>' />
                                <asp:Label ID="lblIUD" runat="server" Text='<%# Eval("IUD") %>' />
                                <asp:Label ID="lblDescricao" runat="server" Text='<%# Eval("Descricao") %>' />
                                <%# DataBinder.Eval(Container.DataItem, "IUD") %>
                                <%# DataBinder.Eval(Container.DataItem, "Descricao") %>
                            </div>
                            <asp:Panel ID="panArquivo" runat="server">
                                <asp:ImageButton runat="server" ID="imbDoc" Style="cursor: pointer; border: 0 none;"
                                    ImageAlign="AbsMiddle" OnClick="imbDoc_Click" />
                                <asp:LinkButton ID="lnkExcluir" Font-Size="10px" runat="server" ForeColor="Black"
                                    Text="Excluir" data-ToolTip="default" ToolTip="Excluir Arquivo" OnClick="lnkExcluir_Click" OnClientClick="if(!confirm('Deseja realmente excluir o arquivo?')) return false;" />
                            </asp:Panel>
                        </ItemTemplate>
                    </asp:DataList>
                </td>
            </tr>
        </table>
        <asp:Label ID="lblErroMsg" ForeColor="red" runat="server" />
    </ContentTemplate>
</asp:UpdatePanel>
