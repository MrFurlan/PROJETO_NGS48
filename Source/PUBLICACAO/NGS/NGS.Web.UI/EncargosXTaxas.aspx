<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="EncargosXTaxas.aspx.vb" Inherits="NGS.Web.UI.EncargosXTaxas" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <script type="text/javascript" src="../Scripts/gridviewScroll.min.js"></script>
    <script type="text/javascript">
        $(document).ready(function () {
            gridviewScroll();
        });

        var w = Gridencargosxtaxas.width;
        alert(w.val);
        function gridviewScroll() {
            $('#<%=Gridencargosxtaxas.ClientID%>').gridviewScroll({
                width: 950,
                height: 385
            });
        }
    </script>

</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngEncargosXTaxas" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlEncargosXTaxas" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Encargos X Taxas
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
                            <asp:LinkButton class="iconLimpar" ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconRelatorio" ID="lnkRelatorio" Text="Relatório" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    UF:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbEstado" runat="server" DataMember="Estados" DataTextField="Descricao"
                        DataValueField="Estado_id" TabIndex="10" Width="600px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Encargo:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="CmbEncargo" runat="server" DataMember="Encargos" DataTextField="Descricao"
                        DataValueField="Descricao" TabIndex="10" Width="600px" />
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Grupo:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlGrupoProduto" runat="server" AutoPostBack="True" Width="600px" />
                </div>
                <div class="coltxt">
                    <asp:LinkButton ID="lnkBuscaProduto" runat="server" Height="20px" Width="20px"><asp:Image ID="imgSearch" ImageUrl="~/Images/search.png" runat="server" data-ToolTip="default"
                            ToolTip="Consulta Produto" /></asp:LinkButton>
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Produto:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlProdutos" runat="server" Width="600px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtData" CssClass="calendario" runat="server" Width="90px" data-ToolTip="default"
                        ToolTip="Data inicial do recolhimento do imposto." />
                </div>
                <div class="collbl">
                    Percentual:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPercentual" runat="server" CssClass="txtDecimal6" data-ToolTip="default"
                        ToolTip="Alíquota do encargo." />
                </div>

                <div class="collbl">
                    Simples Nacional:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtSimplesNacional" runat="server" CssClass="txtDecimal6" />
                </div>
            </div>

            <div class="bordagrid">
                <asp:GridView ID="Gridencargosxtaxas" runat="server" AutoGenerateColumns="False"
                    CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="Gridencargosxtaxas_SelectedIndexChanged">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:CommandField SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button"></asp:CommandField>
                        <asp:BoundField DataField="Estado_id" HeaderText="Uf">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Encargo_id" HeaderText="Encargo" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Produto_id" HeaderText="Produto" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Data_id" HeaderText="Data" DataFormatString="{0:dd/MM/yyyy}"
                            HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Percentual" HeaderText="Percentual">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField DataField="SimplesNacional" HeaderText="Simples Nacional">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaProduto ID="ucConsultaProduto" runat="server" />
</asp:Content>
