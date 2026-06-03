<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="TotalizaLotes.aspx.vb" Inherits="NGS.Web.UI.TotalizaLotes" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngTotalizaLotes" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlTotalizaLotes" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Totalizador De Lotes
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconNovo" runat="server">
                            <asp:LinkButton ID="LinkConsultar" Text="Consultar" runat="server" />
                        </li>
                        <li class="iconRelatorio rel" runat="server"><a>Relatório</a>
                            <ul>
                                <li class="iconPdf">
                                    <asp:LinkButton ID="lnkRelatorio" runat="server" Text="Relatório Pdf" />
                                </li>
                                <li class="iconExcel">
                                    <asp:LinkButton ID="lnkExcelDados" runat="server" Text="Excel Dados" />
                                </li>
                            </ul>
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
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUnidade" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlUnidade_SelectedIndexChanged"
                        Width="616px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                    <asp:CheckBox ID="chkConsolidarEmpresa" ToolTip="Unificar o CNPJs da empresa." CssClass="multiselect" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="616px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Periodo:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicial" CssClass="calendario" runat="server" Width="112px"
                        data-ToolTip="default" ToolTip="Data inicial a final para consulta." />
                    &nbsp;a&nbsp;
                    <asp:TextBox ID="txtDataFinal" runat="server" CssClass="calendario" Width="112px"
                        data-ToolTip="default" ToolTip="Data inicial a final para consulta." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Lote:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtLote" runat="server" Width="112px" data-ToolTip="default" ToolTip="Inserir o número do lote desejado." />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="GridView1" runat="server" Width="100%" CellPadding="1" AutoGenerateColumns="False"
                    ForeColor="#333333" GridLines="Horizontal">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:CommandField SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button">
                            <ItemStyle Width="25px"></ItemStyle>
                        </asp:CommandField>
                        <asp:BoundField DataField="Movimento" DataFormatString="{0:dd/MM/yy}" HeaderText="Movto">
                            <ItemStyle Width="50px"></ItemStyle>
                            <HeaderStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Lote" HeaderText="Lote">
                            <HeaderStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="DebitoOficial" DataFormatString="{0:N}" HeaderText="D&#233;bito R$">
                            <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="CreditoOficial" DataFormatString="{0:N}" HeaderText="Cr&#233;dito R$">
                            <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="SaldoOficial" DataFormatString="{0:N}" HeaderText="Saldo R$">
                            <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="DebitoMoeda" DataFormatString="{0:N}" HeaderText="Debito U$">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField DataField="CreditoMoeda" DataFormatString="{0:N}" HeaderText="Cr&#233;dito U$">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField DataField="SaldoMoeda" DataFormatString="{0:N}" HeaderText="Saldo U$">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
