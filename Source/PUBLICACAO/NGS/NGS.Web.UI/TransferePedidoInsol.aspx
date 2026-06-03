<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master" CodeBehind="TransferePedidoInsol.aspx.vb" Inherits="NGS.Web.UI.TransferePedidoInsol" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <%--    <script type="text/javascript">
        function pageLoadTransferePedidoInsol() {
            $("#<%=lnkAtualizar.ClientID%>").hide();

            $("input[type=radio]", "#MainContent_GridTransferePedidoInsol").each(function () {
                $(this).attr("name", "rbTransferePedido");
                $(this).click(function () { $("#<%=lnkAtualizar.ClientID%>").show("drop"); });
            });

        };

        $(document).ready(function () {
            pageLoadTransferePedidoInsol();
            var prmTransferePedidoInsol = Sys.WebForms.PageRequestManager.getInstance();
            prmTransferePedidoInsol.add_endRequest(pageLoadTransferePedidoInsol);
        });
    </script>--%>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngCCeObservacao" runat="server" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlTransferePedidoInsol" runat="server">
        <ContentTemplate>

            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Transfere Pedido
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconAtualizar" ID="lnkAtualizar" runat="server" Text="Atualizar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconConsultar" ID="lnkConsultar" runat="server" Text="Consultar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconLimpar" ID="lnkLimpar" runat="server" Text="Limpar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" runat="server" Text="Ajuda" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAtualizar" ID="lnkAtualizaNotas" runat="server" Text="Atualizar Notas" />
                        </li>
                    </ul>
                </div>
                <div>
                    <table>
                        <tr id="trFile" runat="server" visible="true">
                            <td style="width: 47px; white-space: nowrap; padding: 0;" valign="top">
                                <span class="HeaderSpanSecond">Arquivo:</span>
                            </td>
                            <td>
                                <uc:File ID="ucFile" runat="server" />
                            </td>
                        </tr>
                    </table>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUnidadeDeNegocio" runat="server" Width="633px" AutoPostBack="True"
                        OnSelectedIndexChanged="ddlUnidadeDeNegocio_SelectedIndexChanged" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="633px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                    <asp:TextBox ID="txtCliente" runat="server" Width="593px" Enabled="False" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnCliente" runat="server" Text=" > " CssClass="btn" UseSubmitBehavior="False"
                        OnClick="btnCliente_Click" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Pedido:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPedido" runat="server" Width="100px" CssClass="txtNumerico" />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="GridTransferePedidoInsol" runat="server" Width="100%" AutoGenerateColumns="False"
                    CellPadding="4" ForeColor="#333333" GridLines="None">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:TemplateField>
                            <ItemTemplate>
                                <asp:CheckBox ID="chkPedido" name="rbTransferePedido" runat="server" />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:BoundField DataField="Codigo" HeaderText="Nº.Pedido">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="DataPedido" HeaderText="Data" DataFormatString="{0:dd/MM/yyyy}">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="NomeCliente" HeaderText="Cliente">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
            <div>
                <asp:TextBox ID="txtObservacoesDeEmbarque" runat="server" Width="98%" Height="155px"
                    ForeColor="Blue" Font-Names="Tahoma" TextMode="MultiLine" />
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaPedidos ID="ucConsultaPedidos" runat="server" />
</asp:Content>
