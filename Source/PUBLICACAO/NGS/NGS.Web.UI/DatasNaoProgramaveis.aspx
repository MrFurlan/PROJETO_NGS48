<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="DatasNaoProgramaveis.aspx.vb" Inherits="NGS.Web.UI.DatasNaoProgramaveis" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngDatasNaoProgramaveis" runat="server" EnableScriptGlobalization="true"
        EnableScriptLocalization="true" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlDatasNaoProgramaveis" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Datas Não Programáveis
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
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlEmpresa" runat="server" Width="581px" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtData" runat="server" ClientIDMode="Static" CssClass="calendario"
                        MaxLength="10" data-ToolTip="default" ToolTip="Feriados e datas especiais para não programar pagamentos."
                        TabIndex="5" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Descrição:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDescricao" runat="server" CssClass="texto" Width="572px" data-ToolTip="default"
                        ToolTip="Descrição do feriado/data especial." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Executar em:
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkTodasFiliais" runat="server" data-ToolTip="default" Text="Todas as filiais" Checked="true"
                        ToolTip="Se checado, será executado em todas as filiais." />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="grd" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="grd_SelectedIndexChanged">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True"
                            ItemStyle-HorizontalAlign="Center" ItemStyle-VerticalAlign="Middle" />
                        <asp:BoundField DataField="Empresa_Id" HeaderText="Empresa" HeaderStyle-HorizontalAlign="Left"
                            ItemStyle-VerticalAlign="Middle" />
                        <asp:BoundField DataField="CidadeUF" HeaderText="Cidade/UF" HeaderStyle-HorizontalAlign="Left"
                            ItemStyle-VerticalAlign="Middle" />
                        <asp:BoundField DataField="Data_Id" HeaderText="Data" HeaderStyle-HorizontalAlign="Left"
                            ItemStyle-VerticalAlign="Middle" />
                        <asp:BoundField DataField="Descricao" HeaderText="Descrição" HeaderStyle-HorizontalAlign="Left"
                            ItemStyle-VerticalAlign="Middle" />
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaEmpresas ID="ucConsultaEmpresas" runat="server" />
</asp:Content>
