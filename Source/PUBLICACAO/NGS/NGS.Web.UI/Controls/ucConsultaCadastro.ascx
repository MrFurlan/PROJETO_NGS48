<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucConsultaCadastro.ascx.vb"
    Inherits="NGS.Web.UI.ucConsultaCadastro" %>
<script type="text/javascript">
    function pageLoadConsultaCadastro() {
    }

    $(document).ready(function () {
        pageLoadConsultaCadastro();
    });

    var prmConsultaCadastro = Sys.WebForms.PageRequestManager.getInstance();
    prmConsultaCadastro.add_endRequest(pageLoadConsultaCadastro);
</script>
<div id="divConsultaCadastro" class="uc" title="Consulta Cadastro" style="display: none;">
    <asp:UpdatePanel ID="updpnlConsultaCadastro" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconConsultar" runat="server">
                            <asp:LinkButton ID="lnkConsultar" Text="Consultar" runat="server" />
                        </li>
                        <li class="iconFechar" runat="server">
                            <asp:LinkButton ID="lnkFechar" Text="Fechar" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div id="rowCnpj" runat="server" class="row">
                <div class="collbluc">
                    CPF/CNPJ:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCpfCnpj" runat="server" Width="200px" CssClass="txtNumerico" />
                </div>
                <div class="collbluc">
                    Uf:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUfConsulta" runat="server" AutoPostBack="True" ToolTip="UF do CPF ou CNPJ a ser utilizada para consulta do cadastro junto a sefaz." />
                </div>
            </div>

            <div class="bordagrid" style="height: 200px;">
                <asp:GridView ID="gridCadastro" runat="server" Width="100%" ForeColor="#333333" GridLines="None"
                    CellPadding="4" AutoGenerateColumns="False">
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White"></FooterStyle>
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333"></RowStyle>
                    <PagerStyle HorizontalAlign="Center" BackColor="#284775" ForeColor="White"></PagerStyle>
                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333"></SelectedRowStyle>
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White"></HeaderStyle>
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775"></AlternatingRowStyle>
                    <EditRowStyle BackColor="#999999"></EditRowStyle>
                    <Columns>
                        <asp:BoundField DataField="0" HeaderText="Conta">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>

                        <asp:TemplateField HeaderText="Cliente">
                            <ItemTemplate>
                                <asp:Label ID="lblCliente" runat="server" Text='<%# Eval("1")%>' />
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:TemplateField>


                        <asp:TemplateField HeaderText="Descrição">
                            <ItemTemplate>
                                <asp:Label ID="lblDescricao" runat="server" Text='<%# Eval("2")%>' />
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:TemplateField>

                        <asp:BoundField DataField="3" HeaderText="Movimento" HtmlEncode="False" DataFormatString="{0:dd/MM/yyyy}">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>

                      <%--  <asp:BoundField DataField="Lote" HeaderText="Lote">
                            <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                        </asp:BoundField>

                        <asp:BoundField DataField="CodigoProduto" HeaderText="Produto">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>

                        <asp:BoundField DataField="CodigoCusto" HeaderText="Custo">
                            <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                        </asp:BoundField>

                        <asp:BoundField DataField="Historico" HeaderText="Histórico">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>

                        <asp:BoundField DataField="DebitoOficial" DataFormatString="{0:N2}" HeaderText="D&#233;bito">
                            <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                        </asp:BoundField>

                        <asp:BoundField DataField="CreditoOficial" DataFormatString="{0:N2}" HeaderText="Cr&#233;dito">
                            <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                        </asp:BoundField>

                        <asp:BoundField DataField="Saldo" DataFormatString="{0:N2}" HeaderText="Saldo">
                            <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                        </asp:BoundField>--%>

                    </Columns>
                </asp:GridView>
            </div>



            <%-- <div class="coltxt">
                <asp:ImageButton ID="imgConsultaCadastro" runat="server" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png" ImageAlign="AbsMiddle"
                    CssClass="imgconsultar" data-ToolTip="default" ToolTip="Consulta o cadastro na Sefaz." />
            </div>--%>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
