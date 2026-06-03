<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="BalancoAuditadoMes.aspx.vb" Inherits="NGS.Web.UI.BalancoAuditadoMes" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngBalanco" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="UpdateBalanco" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="Hid" runat="server" />
            <div class="titulodiv">
                Balanço Auditado Fechamento Mês a Mês
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlAno" runat="server" Width="150px" AutoPostBack="True" Style="margin-left: 4px;"
                        data-ToolTip="default" ToolTip="Selecionar o ano de consulta/auditoria." />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="gridMeses" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:BoundField DataField="DescricaoEmpresa" HeaderText="Descricao">
                            <ControlStyle Width="250px" />
                            <FooterStyle Width="250px" />
                            <HeaderStyle Width="250px" />
                            <ItemStyle Width="250px" />
                        </asp:BoundField>
                        <asp:TemplateField HeaderText="Janeiro">
                            <ItemTemplate>
                                <asp:Button ID="BtnJaneiro" runat="server" Text="Auditado" Font-Bold="True" OnClick="Btn_Click"
                                    Width="56px" />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Fevereiro">
                            <ItemTemplate>
                                <asp:Button ID="BtnFevereiro" runat="server" Text="Auditado" Font-Bold="True" OnClick="Btn_Click"
                                    Width="56px" />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Março">
                            <ItemTemplate>
                                <asp:Button ID="BtnMarco" runat="server" Text="Auditado" Font-Bold="True" OnClick="Btn_Click"
                                    Width="56px" />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Abril">
                            <ItemTemplate>
                                <asp:Button ID="BtnAbril" runat="server" Text="Auditado" Font-Bold="True" OnClick="Btn_Click"
                                    Width="56px" />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Maio">
                            <ItemTemplate>
                                <asp:Button ID="BtnMaio" runat="server" Text="Auditado" Font-Bold="True" OnClick="Btn_Click"
                                    Width="56px" />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Junho">
                            <ItemTemplate>
                                <asp:Button ID="BtnJunho" runat="server" Text="Auditado" Font-Bold="True" OnClick="Btn_Click"
                                    Width="56px" />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Julho">
                            <ItemTemplate>
                                <asp:Button ID="BtnJulho" runat="server" Text="Auditado" Font-Bold="True" OnClick="Btn_Click"
                                    Width="56px" />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Agosto">
                            <ItemTemplate>
                                <asp:Button ID="BtnAgosto" runat="server" Text="Auditado" Font-Bold="True" OnClick="Btn_Click"
                                    Width="56px" />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Setembro">
                            <ItemTemplate>
                                <asp:Button ID="BtnSetembro" runat="server" Text="Auditado" Font-Bold="True" OnClick="Btn_Click"
                                    Width="56px" />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Outubro">
                            <ItemTemplate>
                                <asp:Button ID="BtnOutubro" runat="server" Text="Auditado" Font-Bold="True" OnClick="Btn_Click"
                                    Width="56px" />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Novembro">
                            <ItemTemplate>
                                <asp:Button ID="BtnNovembro" runat="server" Text="Auditado" Font-Bold="True" OnClick="Btn_Click"
                                    Width="56px" />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Dezembro">
                            <ItemTemplate>
                                <asp:Button ID="BtnDezembro" runat="server" Text="Auditado" Font-Bold="True" OnClick="Btn_Click"
                                    Width="56px" />
                            </ItemTemplate>
                        </asp:TemplateField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
