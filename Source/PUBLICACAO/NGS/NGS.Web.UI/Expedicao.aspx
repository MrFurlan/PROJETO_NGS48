<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="Expedicao.aspx.vb" Inherits="NGS.Web.UI.Expedicao" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrm" runat="server" />
    <table width="100%" id="tblNotasConferencia" runat="server" visible="false">
        <tr>
            <td class="subtitulodiv">
                <label>
                    Notas para conferência
                </label>
            </td>
        </tr>
        <tr>
            <td>
                <div class="bordagrid">
                    <asp:GridView ID="gdvConferenciaNota" runat="server" AutoGenerateColumns="False"
                        CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                        <Columns>
                            <asp:BoundField DataField="Codigo" HeaderText="Nota" />
                            <asp:BoundField DataField="Serie" HeaderText="Série" />
                            <asp:BoundField DataField="CodigoEmpresa" HeaderText="Empresa" ItemStyle-CssClass="none"
                                HeaderStyle-CssClass="none" />
                            <asp:BoundField DataField="EnderecoEmpresa" ItemStyle-CssClass="none" HeaderStyle-CssClass="none" />
                            <asp:BoundField DataField="Empresa.Cidade" HeaderText="Cidade" ItemStyle-CssClass="none"
                                HeaderStyle-CssClass="none" />
                            <asp:BoundField DataField="CodigoCliente" ItemStyle-CssClass="none" HeaderStyle-CssClass="none" />
                            <asp:BoundField DataField="EnderecoCliente" ItemStyle-CssClass="none" HeaderStyle-CssClass="none" />
                            <asp:BoundField DataField="Cliente.Nome" HeaderText="Fornecedor" ItemStyle-CssClass="none"
                                HeaderStyle-CssClass="none" />
                            <asp:TemplateField HeaderText="Empresa">
                                <ItemTemplate>
                                    <%# Convert.ToDouble(DataBinder.Eval(Container.DataItem,"CodigoEmpresa")).ToString("000\.000\.000\/0000\-00") %>
                                    -
                                    <%# DataBinder.Eval(Container.DataItem,"Empresa.Nome") %><br />
                                    <%# DataBinder.Eval(Container.DataItem,"Empresa.Cidade") %>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Fornecedor">
                                <ItemTemplate>
                                    <%# Convert.ToDouble(DataBinder.Eval(Container.DataItem, "CodigoCliente")).ToString("000\.000\.000\/0000\-00")%>
                                    -
                                    <%# DataBinder.Eval(Container.DataItem, "Cliente.Nome")%><br />
                                    <%# DataBinder.Eval(Container.DataItem,"Cliente.Cidade") %>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField>
                                <ItemTemplate>
                                    <asp:LinkButton ID="lnkConferir" runat="server" Text="Conferir" OnClick="lnkConferir_Click"
                                        ForeColor="Black"></asp:LinkButton>
                                </ItemTemplate>
                            </asp:TemplateField>
                        </Columns>
                        <EditRowStyle BackColor="#999999" />
                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                        <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    </asp:GridView>
                </div>
            </td>
        </tr>
    </table>
</asp:Content>
