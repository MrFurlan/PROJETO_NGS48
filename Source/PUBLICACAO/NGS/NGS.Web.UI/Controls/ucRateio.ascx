<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucRateio.ascx.vb" Inherits="NGS.Web.UI.ucRateio" %>
<div id="divRateio" class="uc" title="Informe os dados de rateio" style="display: none;">
    <asp:UpdatePanel ID="updpnlRateio" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <asp:HiddenField ID="hdfIndex" runat="server" />
            <asp:HiddenField ID="hdfEdicao" runat="server" />
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconMais" runat="server">
                            <asp:LinkButton ID="lnkAdicionar" Text="Adicionar" runat="server" />
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li class="iconFechar" runat="server">
                            <asp:LinkButton ID="lnkFechar" Text="Fechar" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="painelleft">
                <div class="row">
                    <div class="collbluc">
                        Produto
                    </div>
                    <div class="coltxt">
                        <asp:Label ID="lblProduto" runat="server" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc">
                        Unidade:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlUnidadeNegocio" runat="server" AutoPostBack="True" Width="600px" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc">
                        Empresa:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlEmpresa" runat="server" AutoPostBack="True" Width="600px" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc">
                        Centro de Custo:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlCentroCusto" runat="server" Width="600px" />
                    </div>
                </div>
            </div>
            <div class="painelright">
                <div class="row">
                    <div class="collbluc">
                        Valor Total:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtTotal" runat="server" CssClass="txtDecimal" Enabled="false" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc">
                        Valor Rateado:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtValorRateado" runat="server" CssClass="txtDecimal" Enabled="false" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc">
                        Saldo Rateio:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtSaldoRateio" runat="server" CssClass="txtDecimal" Enabled="false" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc">
                        Valor:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtValorRateio" runat="server" Width="100px" CssClass="txtDecimal" />
                    </div>
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="gridRateio" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    Width="100%" ForeColor="#333333" GridLines="None" OnSelectedIndexChanged="gridRateio_SelectedIndexChanged" DataKeyNames="CodigoCentroDeCusto">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                        <asp:TemplateField HeaderText="Un. Negócio">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                            <ItemTemplate>
                                <asp:Label ID="lblUnidadeDeNegocio" runat="server" Text='<%# Eval("CodigoUnidadeDeNegocio") & " - " & Eval("NomeUnidadeDeNegocio")%>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Empresa">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                            <ItemTemplate>
                                <asp:Label ID="lblEmpresa" runat="server" Text='<%# Eval("CodigoEmpresaRateio") & " - " & Eval("EndEmpresaRateio") & " - " & Eval("NomeEmpresa")%>' />
                            </ItemTemplate>
                        </asp:TemplateField>


                        <asp:BoundField DataField="CodigoCentroDeCusto" HeaderText="CC">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>


                        <asp:TemplateField HeaderText="Centro de Custo">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                            <ItemTemplate>
                                <asp:Label ID="lblCentroDeCusto" runat="server" Text='<%# Eval("CodigoCentroDeCusto") & " - " & Eval("CentroDeCusto.Descricao")%>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:BoundField DataField="Valor" DataFormatString="{0:N2}" HeaderText="Valor">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:TemplateField HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                            <ItemTemplate>
                                <asp:LinkButton ID="lnkRemover" runat="server" OnClick="lnkRemover_Click">
                                    <asp:Image ID="imgExcluir" runat="server" Width="16px" Height="16px" ImageUrl="~/Images/deletar.gif"
                                        data-ToolTip="default" ToolTip="Excluir" />
                                </asp:LinkButton>
                            </ItemTemplate>
                        </asp:TemplateField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
