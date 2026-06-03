<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="Arrendante.aspx.vb" Inherits="NGS.Web.UI.Arrendante" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngArrendante" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlArrendante" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Arrendantes
            </div>
            <div class="menu_acoes">
                <div class="acoes">

                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconConsultar" ID="lnkConsultar" Text="Consultar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconNovo" ID="lnkConfirmar" Text="Confirmar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoArrendante" runat="server" />
                    <asp:TextBox ID="txtArrendante" runat="server" Enabled="False" Width="400px" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnArrendante" runat="server" OnClick="btnArrendante_Click" Text=">"
                        CssClass="btn" UseSubmitBehavior="False" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data Contrato:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataContrato" runat="server" Width="86px" CssClass="calendario"
                        AutoPostBack="True" OnTextChanged="txtDataContrato_TextChanged" />
                </div>
                <div class="collbl">
                    Data Vencimento:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataVencimento" runat="server" Width="86px" CssClass="calendario"
                        AutoPostBack="True" OnTextChanged="txtDataVencimento_TextChanged" />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="gridMatricula" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="gridVencimentos_SelectedIndexChanged">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:BoundField DataField="Matricula_Id" HeaderText="Matricula" />
                        <asp:BoundField DataField="Area" HeaderText="Area" />
                        <asp:BoundField DataField="AreaArrendada" HeaderText="Arrendada" />
                        <asp:BoundField DataField="Saldo" HeaderText="Saldo" />
                        <asp:TemplateField HeaderText="Arrendamento">
                            <ItemTemplate>
                                <asp:TextBox ID="txtArrendamento" runat="server" onkeyup="validar_valor(this, 15);"
                                    value="0,00" />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Observacao">
                            <ItemTemplate>
                                <asp:TextBox ID="txtObservacao" runat="server" Width="557px" />
                            </ItemTemplate>
                        </asp:TemplateField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
