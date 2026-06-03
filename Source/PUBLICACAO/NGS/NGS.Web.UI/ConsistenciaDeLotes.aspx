<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="ConsistenciaDeLotes.aspx.vb" Inherits="NGS.Web.UI.ConsistenciaDeLotes" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngConsistenciaDeLotes" runat="server" AsyncPostBackTimeout="900" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlConsistenciaDeLotes" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Consistência de Lotes
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server" class="iconConsultar">
                            <asp:LinkButton ID="LinkConsultar" runat="server" Text="Consultar" />
                        </li>
                        <li class="iconRelatorio rel" runat="server"><a>Relatório</a>
                            <ul>
                                <li class="iconPdf">
                                    <asp:LinkButton ID="lnkRelatorio" runat="server" Text="Pdf" />
                                </li>
                                <li class="iconExcel">
                                    <asp:LinkButton ID="lnkExcelDados" runat="server" Text="Excel" />
                                </li>
                            </ul>
                        </li>
                        <li runat="server" class="iconLimpar">
                            <asp:LinkButton ID="lnkLimpar" runat="server" Text="Limpar" />
                        </li>
                        <li runat="server" class="iconAjuda">
                            <asp:LinkButton ID="lnkAjuda" runat="server" Text="Ajuda" />
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
                </div>
                <div class="collbl">
                    à:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" runat="server" CssClass="calendario" Width="112px"
                        data-ToolTip="default" ToolTip="Data inicial a final para consulta." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Lote:
                </div>
                <div class="coltxt" style="width: 142px">
                    <asp:TextBox ID="txtLote" runat="server" Width="112px" data-ToolTip="default"
                        ToolTip="Inserir o número do lote desejado." />
                </div>
                <div class="collbl">
                    Título:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtTitulo" runat="server" Width="112px" data-ToolTip="default"
                        ToolTip="Inserir o número do título desejado." />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="gridConsistenciaDeLotes" runat="server" AutoGenerateColumns="False"
                    CellPadding="4" ForeColor="#333333" Width="100%" GridLines="None">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:CommandField SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button">
                            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle"></ItemStyle>
                        </asp:CommandField>
                        <asp:BoundField DataField="Empresa" HeaderText="Empresa">
                            <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                            <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Movimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Movto">
                            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Lote" HeaderText="Lote"></asp:BoundField>
                        <asp:BoundField DataField="Sequencia" HeaderText="Seq."></asp:BoundField>
                        <asp:BoundField DataField="Conta" HeaderText="Conta">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Reduzido" HeaderText="Cliente"></asp:BoundField>
                        <asp:BoundField DataField="CUSTO" HeaderText="C.C"></asp:BoundField>
                        <asp:BoundField DataField="HISTORICO" HeaderText="Hist&#243;rico">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="DEBITO" DataFormatString="{0:N2}" HeaderText="D&#233;bito">
                            <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="CREDITO" DataFormatString="{0:N2}" HeaderText="Cr&#233;dito">
                            <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="SALDO" DataFormatString="{0:N2}" HeaderText="Saldo">
                            <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
