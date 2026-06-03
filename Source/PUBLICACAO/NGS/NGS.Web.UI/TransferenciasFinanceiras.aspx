<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="TransferenciasFinanceiras.aspx.vb" Inherits="NGS.Web.UI.TransferenciasFinanceiras" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmTransferenciasFinanceiras" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True">
    </asp:ScriptManager>
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlTransferenciasFinanceiras" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Transferências Financeiras
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconNovo" runat="server">
                            <asp:LinkButton ID="lnkNovo" Text="Gravar" runat="server" />
                        </li>
                        <li class="iconAtualizar" runat="server">
                            <asp:LinkButton ID="lnkAtualizar" Text="Atualizar" runat="server" />
                        </li>
                        <li class="iconExcluir" runat="server">
                            <asp:LinkButton ID="lnkExcluir" Text="Excluir" runat="server" OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" />
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li class="iconRelatorio" runat="server">
                            <asp:LinkButton ID="lnkRelatorio" Text="Relatório" runat="server" />
                        </li>
                        <li class="iconAjuda" runat="server">
                            <asp:LinkButton ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa Débito:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="hdfEmpresaDebito" runat="server" />
                    <asp:TextBox ID="txtEmpresaDebito" runat="server" CssClass="texto" Width="500px"
                        data-ToolTip="default" ToolTip="Informar a empresa que realizará o débito." />
                </div>
                <div class="coltxt">
                    <asp:ImageButton ID="btnMostraEmpresaDebito" runat="server" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png"
                        BorderStyle="None" OnClick="btnMostraEmpresaDebito_Click" Width="19px" Height="19px"
                        data-ToolTip="default" ToolTip="Procurar a empresa que realizará o débito." />
                </div>
                <div class="coltxt">
                    <asp:ImageButton ID="btnLimpaEmpresaDebito" runat="server" ImageUrl="~/Images/Borracha.jpg"
                        CssClass="btn" BorderStyle="None" OnClick="btnLimpaEmpresaDebito_Click" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa Crédito:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="hdfEmpresaCredito" runat="server" />
                    <asp:TextBox ID="txtEmpresaCredito" runat="server" CssClass="texto" Width="500px"
                        data-ToolTip="default" ToolTip="Informar a empresa que realizará o crédito." />
                </div>
                <div class="coltxt">
                    <asp:ImageButton ID="btnMostraEmpresaCredito" runat="server" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png"
                        Width="19px" Height="19px" BorderStyle="None" OnClick="btnMostraEmpresaCredito_Click"
                        data-ToolTip="default" ToolTip="Procurar a empresa que realizará o crédito." />
                </div>
                <div class="coltxt">
                    <asp:ImageButton ID="btnLimpaEmpresaCredito" runat="server" ImageUrl="~/Images/Borracha.jpg"
                        CssClass="btn" BorderStyle="None" OnClick="btnLimpaEmpresaCredito_Click" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa Contábil:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="hdfEmpresaContabil" runat="server" />
                    <asp:TextBox ID="txtEmpresaContabil" runat="server" CssClass="texto" Width="500px"
                        data-ToolTip="default" ToolTip="Empresa que está transferindo." />
                </div>
                <div class="coltxt">
                    <asp:ImageButton ID="btnMostraEmpresaContabil" runat="server" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png"
                        Width="19px" Height="19px" BorderStyle="None" OnClick="btnMostraEmpresaContabil_Click"
                        data-ToolTip="default" ToolTip="Procurar empresa que está transferindo." />
                </div>
                <div class="coltxt">
                    <asp:ImageButton ID="btnLimpaEmpresaContabil" runat="server" ImageUrl="~/Images/Borracha.jpg"
                        CssClass="btn" BorderStyle="None" OnClick="btnLimpaEmpresaContabil_Click" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Conta Contábil:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlContaContabil" runat="server" CssClass="texto" Width="504px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente Contábil:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="hdfClienteContabil" runat="server" />
                    <asp:TextBox ID="txtClienteContabil" runat="server" CssClass="texto" Width="500px"
                        data-ToolTip="default" ToolTip="Empresa que receberá a transferência." />
                </div>
                <div class="coltxt">
                    <asp:ImageButton ID="btnMostraClienteContabil" runat="server" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png"
                        Width="19px" Height="19px" BorderStyle="None" OnClick="btnMostraClienteContabil_Click"
                        data-ToolTip="default" ToolTip="Procurar empresa que receberá a transferência." />
                </div>
                <div class="coltxt">
                    <asp:ImageButton ID="btnLimpaClienteContabil" runat="server" ImageUrl="~/Images/Borracha.jpg"
                        CssClass="btn" BorderStyle="None" OnClick="btnLimpaClienteContabil_Click" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Tipo de Operação:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdoDebito" runat="server" Text="Dédito" GroupName="grdTipoDeOperacao"
                        data-ToolTip="default" ToolTip="Informar se é uma operação a débito ou a crédito." />
                    <asp:RadioButton ID="rdoCredito" runat="server" Text="Crédito" Checked="true" GroupName="grdTipoDeOperacao"
                        data-ToolTip="default" ToolTip="Informar se é uma operação a débito ou a crédito." />
                </div>
                <div class="collbl">
                    Opção:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdoEmiteAvisoA" runat="server" Text="Emite Aviso" data-ToolTip="default"
                        ToolTip="Selecionar para emitir, ou não, aviso." />
                    <asp:RadioButton ID="rdoEmiteAvisoB" runat="server" Text="Não Emite Aviso" Checked="true"
                        data-ToolTip="default" ToolTip="Selecionar para emitir, ou não, aviso." />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="grd" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="grd_SelectedIndexChanged">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True"
                            ItemStyle-HorizontalAlign="Center" ItemStyle-VerticalAlign="Middle" />
                        <asp:BoundField DataField="EmpresaDebito" HeaderText="Emp. Débito" HeaderStyle-HorizontalAlign="Left"
                            ItemStyle-HorizontalAlign="Left" />
                        <asp:BoundField DataField="EnderecoDebito" HeaderText="End." HeaderStyle-HorizontalAlign="Left"
                            ItemStyle-HorizontalAlign="Left" />
                        <asp:BoundField DataField="EmpresaCredito" HeaderText="Emp. Crédito" HeaderStyle-HorizontalAlign="Left"
                            ItemStyle-HorizontalAlign="Left" />
                        <asp:BoundField DataField="EnderecoCredito" HeaderText="End." HeaderStyle-HorizontalAlign="Left"
                            ItemStyle-HorizontalAlign="Left" />
                        <asp:BoundField DataField="EmpresaContabil" HeaderText="Emp. Contábil" HeaderStyle-HorizontalAlign="Left"
                            ItemStyle-HorizontalAlign="Left" />
                        <asp:BoundField DataField="EnderecoContabil" HeaderText="End." HeaderStyle-HorizontalAlign="Left"
                            ItemStyle-HorizontalAlign="Left" />
                        <asp:BoundField DataField="ContaContabil" HeaderText="Conta Contábil" HeaderStyle-HorizontalAlign="Left"
                            ItemStyle-HorizontalAlign="Left" />
                        <asp:BoundField DataField="ClienteContabil" HeaderText="Cliente Contábil" HeaderStyle-HorizontalAlign="Left"
                            ItemStyle-HorizontalAlign="Left" />
                        <asp:BoundField DataField="EndClienteContabil" HeaderText="End." HeaderStyle-HorizontalAlign="Left"
                            ItemStyle-HorizontalAlign="Left" />
                        <asp:BoundField DataField="DebitoCredito" HeaderText="D/C" HeaderStyle-HorizontalAlign="Left"
                            ItemStyle-HorizontalAlign="Left" />
                        <asp:BoundField DataField="EmiteAviso" HeaderText="Aviso" HeaderStyle-HorizontalAlign="Left"
                            ItemStyle-HorizontalAlign="Left" />
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaEmpresas ID="ucConsultaEmpresas" runat="server" />
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
