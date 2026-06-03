<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="AcompanhamentoDeEntregaDeProduto.aspx.vb" Inherits="NGS.Web.UI.AcompanhamentoDeEntregaDeProduto" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <script type="text/javascript">
        function pageLoadRelatorioEntradaSaidaNotas() {
            $("#MainContent_lstClasses").multiselect({
                header: "Escolha os encargos!",
                selectedList: 4
            }).multiselectfilter();
        }

        $(document).ready(function () {
            pageLoadRelatorioEntradaSaidaNotas();
            var prmRelatorioEntradaSaidaNotas = Sys.WebForms.PageRequestManager.getInstance();
            prmRelatorioEntradaSaidaNotas.add_endRequest(pageLoadRelatorioEntradaSaidaNotas);
        });
    </script>

    <style type="text/css">
        .collbl {
            width: 170px;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmgAcphntoDeProduto" runat="server" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlAcphntoDeProduto" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Acompanhamento de entrega de produto - Retenção Clientes
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server" class="iconConsultar">
                            <asp:LinkButton ID="lnkConsultar" runat="server" Text="Consultar" />
                        </li>
                        <li class="iconRelatorio rel" runat="server"><a>Relatório</a>
                            <ul>
                                <li>
                                    <asp:LinkButton class="iconPdf" ID="lnkPdf" runat="server" Text="Pdf" />
                                </li>
                                <li>
                                    <asp:LinkButton class="iconExcel" ID="lnkExcel" runat="server" Text="Excel" />
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
                    <asp:DropDownList ID="ddlUnidade" runat="server" Width="595px" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkConsolidarEmpresa" runat="server"
                        Text="Empresa:" data-ToolTip="default" ToolTip="Consolidar o cnpj da empresa." />
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="595px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Classes de Operações:
                </div>
                <div class="coltxt">
                    <asp:ListBox ID="lstClasses" runat="server" CssClass="multiselect" SelectionMode="Multiple"
                        Width="595px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Produto:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNomeProduto" Enabled="False" runat="server" Width="555px" />
                    <asp:HiddenField ID="txtCodigoProduto" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnProduto" runat="server" CssClass="btn" Text=">" data-ToolTip="default"
                        ToolTip="Seleciona o produto desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Com Movimento Nos Ultimos:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtMovimentoDia" CssClass="txtNumerico" runat="server" Width="80px"
                        data-ToolTip="default" ToolTip="Informar o número de dias." />
                </div>
                <div class="coltxt" style="width: 30px;">
                    Dias
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Entregue a:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtUltimoDiaMovimentoNota" runat="server" CssClass="txtNumerico"
                        Width="80px" data-ToolTip="default" ToolTip="Informar o número de dias." />
                </div>
                <div class="coltxt">
                    Dias
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="grdRetencao" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:BoundField DataField="Empresa_Id" HeaderText="Empresa">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="NomeEmpresa" HeaderText="Nome Empresa">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Cliente" HeaderText="Cliente">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Nome" HeaderText="Nome Cliente">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="UltimaCarga" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Ult. Laudo">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="QtdeFisicaPeriodoTotal" DataFormatString="{0:N0}" HeaderText="Qtde Física">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaProduto ID="ucConsultaProduto" runat="server" />
</asp:Content>
