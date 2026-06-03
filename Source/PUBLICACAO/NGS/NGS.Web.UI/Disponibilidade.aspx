<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="Disponibilidade.aspx.vb" Inherits="NGS.Web.UI.Disponibilidade" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <script type="text/javascript">
        function pageLoadDisponibilidade() {
            $("div.accordion").accordion({
                active: false,
                collapsible: true,
                alwaysOpen: false,
                heightStyle: "content",
                autoHeight: false,
                clearStyle: true
            });
        }

        $(document).ready(function () {
            pageLoadDisponibilidade();
            var prmDisponibilidade = Sys.WebForms.PageRequestManager.getInstance();
            prmDisponibilidade.add_endRequest(pageLoadDisponibilidade);
        });
    </script>
    <style type="text/css">
        .collbl {
            width: 150px;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngDisponibilidade" runat="server" AsyncPostBackTimeout="5000" />
    <asp:HiddenField ID="HID" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlDisponibilidade" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Disponibilidade
            </div>
            <ajaxToolkit:TabContainer ID="tcDisponibilidade" runat="server" ActiveTabIndex="0"
                Width="100%">
                <ajaxToolkit:TabPanel ID="TabPanel1" runat="server" HeaderText="TabPanel1">
                    <HeaderTemplate>
                        Disponibilidade
                    </HeaderTemplate>
                    <ContentTemplate>
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
                                <asp:DropDownList ID="ddlEmpresa" runat="server" AutoPostBack="True" Width="618px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Disponibilidade:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDataDeEmissao" runat="server" CssClass="calendario" Width="110px" data-ToolTip="default" ToolTip="Inserir data da disponibilidade." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Safra Inicial:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlSafraDe" runat="server" Width="618px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Safra Final:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlSafraAte" runat="server" Width="617px" />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="gridDisponibilidade" Width="100%" runat="server" AutoGenerateColumns="False"
                                CellPadding="4" ForeColor="#333333" GridLines="None" CssClass="bordasimples">
                                <Columns>
                                    <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                                    <asp:TemplateField HeaderText="Rel.">
                                        <ItemTemplate>
                                            <asp:LinkButton ID="lnkPdf" runat="server" OnClick="lnkPdf_Click">
                                                <asp:Image ID="imgpdf" runat="server" Width="24px" Height="24px" ImageUrl="~/Images/pdf.png" />
                                            </asp:LinkButton>
                                            <asp:LinkButton ID="lnkExcel" runat="server" OnClick="lnkExcel_Click">
                                                <asp:Image ID="imgexcel" runat="server" Width="24px" Height="24px" ImageUrl="~/Images/excel.png" />
                                            </asp:LinkButton>
                                        </ItemTemplate>
                                        <ItemStyle Width="70px" />
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="CodigoDisponibilidade" HeaderText="Disp." />
                                    <asp:CheckBoxField DataField="Consolidado" HeaderText="Cons." />
                                    <asp:BoundField DataField="CodigoEmpresa" HeaderText="Empresa" />
                                    <asp:BoundField DataField="EndEmpresa" HeaderText="End." />
                                    <asp:BoundField DataField="NomeEmpresa" HeaderText="Nome" />
                                    <asp:BoundField DataField="DataInicial" DataFormatString="{0:d}" HeaderText="Inicio Apuracao Em">
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="DataInicialCargaPatio" DataFormatString="{0:d}" HeaderText="Inicio Carga Patio Em" />
                                    <asp:BoundField DataField="DescricaoProdutos" HeaderText="Produtos" />
                                    <asp:TemplateField HeaderText="Excluir">
                                        <ItemTemplate>
                                            <asp:LinkButton ID="lnkExcluirDisponibilidade" runat="server" OnClientClick=" return confirm('Deseja excluir o registro?');"
                                                OnClick="lnkExcluirDisponibilidade_Click">
                                                <asp:Image ID="imgExcluirDisponibilidade" runat="server" Width="16px" Height="16px"
                                                    ImageUrl="~/Images/excluir16x16.png" data-ToolTip="default" ToolTip="Excluir" />
                                            </asp:LinkButton>
                                        </ItemTemplate>
                                        <ItemStyle HorizontalAlign="Center" />
                                    </asp:TemplateField>
                                </Columns>
                                <EditRowStyle BackColor="#999999" />
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                <SortedAscendingCellStyle BackColor="#E9E7E2" />
                                <SortedAscendingHeaderStyle BackColor="#506C8C" />
                                <SortedDescendingCellStyle BackColor="#FFFDF8" />
                                <SortedDescendingHeaderStyle BackColor="#6F8DAE" />
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel2" runat="server" HeaderText="TabPanel2">
                    <HeaderTemplate>
                        Configuração
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server" style="width: 18%;">
                                        <asp:LinkButton class="iconNovo" ID="lnkNovo" Text="Nova Configuração" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAjuda" ID="lnkAjudaConf" Text="Ajuda" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                <asp:CheckBox ID="chkConsolidarEmpresa" ToolTip="Consolidar o cnpj da empresa." runat="server" Text="Consolidar Empresa:" />
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlEmpresaConfiguracao" runat="server" Width="618px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Data Início da Apuração:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDataInicioDisponibilidade" runat="server" CssClass="calendario" data-ToolTip="default" ToolTip="Informar a data inicial da apuração." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Data Início Carga No Patio:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDataCargaPatio" runat="server" CssClass="calendario" data-ToolTip="default" ToolTip="Data que teve início da carga no pátio." />
                            </div>
                        </div>
                        <div class="row">
                            <div id="divDeposito" class="accordion">
                                <h3>
                                    <asp:Label ID="lblNome" runat="server" Text="Saldos Iniciais Depósitos" />
                                </h3>
                                <div style="height: 100%;">
                                    <div class="row">
                                        <div class="collbluc">
                                            Depósito:
                                        </div>
                                        <div class="coltxt">
                                            <asp:HiddenField ID="hdfCodigoCliente" runat="server" ClientIDMode="Static" />
                                            <asp:TextBox ID="txtNomeCliente" runat="server" ClientIDMode="Static" Enabled="False"
                                                Width="500px" />
                                        </div>
                                        <div class="coltxt">
                                            <asp:Button ID="btnDeposito" runat="server" Text=">" CssClass="btn" data-ToolTip="default" ToolTip="Selecionar o depósito desejado para consulta." />
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="collbluc">
                                            Qtde Fiscal:
                                        </div>
                                        <div class="coltxt">
                                            <asp:TextBox ID="txtQtdeFiscal" runat="server" CssClass="txtDecimal4" Width="120px" data-ToolTip="default" ToolTip="Informar a quantidade fical do estoque." />
                                        </div>
                                        <div class="collbluc">
                                            Qtde Fisica:
                                        </div>
                                        <div class="coltxt">
                                            <asp:TextBox ID="txtQtdeFisica" runat="server" CssClass="txtDecimal4" Width="120px" data-ToolTip="default" ToolTip="Informar a quantidade física do estoque." />
                                        </div>
                                        <div class="coltxt" style="margin-top: 6px;">
                                            <asp:ImageButton ID="imgAdicionarDeposito" runat="server" ImageUrl="~/Images/ico-mais.gif"
                                                Width="16px" />
                                        </div>
                                    </div>
                                    <div class="bordagrid">
                                        <asp:GridView ID="gridDeposito" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                            ForeColor="#333333" GridLines="None" Width="100%">
                                            <Columns>
                                                <asp:BoundField DataField="CodigoDeposito" HeaderText="Depósito">
                                                    <HeaderStyle HorizontalAlign="Left" />
                                                </asp:BoundField>
                                                <asp:BoundField DataField="EndDeposito" HeaderText="End.">
                                                    <HeaderStyle HorizontalAlign="Left" />
                                                </asp:BoundField>
                                                <asp:BoundField DataField="NomeDeposito" HeaderText="Nome">
                                                    <HeaderStyle HorizontalAlign="Left" />
                                                </asp:BoundField>
                                                <asp:BoundField DataField="SaldoInicialFiscal" DataFormatString="{0:N4}" HeaderText="Saldo Fiscal">
                                                    <HeaderStyle HorizontalAlign="Right" />
                                                    <ItemStyle HorizontalAlign="Right" />
                                                </asp:BoundField>
                                                <asp:BoundField DataField="SaldoInicialFisico" DataFormatString="{0:N4}" HeaderText="Saldo Fisico">
                                                    <HeaderStyle HorizontalAlign="Right" />
                                                    <ItemStyle HorizontalAlign="Right" />
                                                </asp:BoundField>
                                                <asp:TemplateField HeaderText="Excluir">
                                                    <ItemTemplate>
                                                        <asp:LinkButton ID="lnkExcluirDeposito" runat="server" OnClientClick=" return confirm('Deseja excluir o registro?');"
                                                            OnClick="lnkExcluirDeposito_Click">
                                                            <asp:Image ID="imgExcluirDisponibilidade" runat="server" Width="16px" Height="16px"
                                                                ImageUrl="~/Images/excluir16x16.png" data-ToolTip="default" ToolTip="Excluir" />
                                                        </asp:LinkButton>
                                                    </ItemTemplate>
                                                    <ItemStyle HorizontalAlign="Center" />
                                                </asp:TemplateField>
                                            </Columns>
                                            <RowStyle BackColor="#EFF3FB" />
                                            <EditRowStyle BackColor="#2461BF" />
                                            <AlternatingRowStyle BackColor="White" />
                                            <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                                            <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                                            <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
                                            <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
                                            <SortedAscendingCellStyle BackColor="#F5F7FB" />
                                            <SortedAscendingHeaderStyle BackColor="#6D95E1" />
                                            <SortedDescendingCellStyle BackColor="#E9EBEF" />
                                            <SortedDescendingHeaderStyle BackColor="#4870BE" />
                                        </asp:GridView>
                                    </div>
                                </div>
                            </div>
                        </div>
                        <div class="row">
                            <uc:SelecaoProduto ID="ucSelecaoProduto" runat="server" />
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaCliente" runat="server" />
</asp:Content>
