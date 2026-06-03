<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="AtivosXContas.aspx.vb" Inherits="NGS.Web.UI.AtivosXContas" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        .collbl {
            width: 145px !important;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmnAtivosXContas" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlAtivosXContas" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                <label>
                    Atixos X Contas
                </label>
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
                            <asp:LinkButton class="iconConsultar" ID="lnkConsultar" Text="Consultar" runat="server" />
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
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlUnidadeNegocio" runat="server" AutoPostBack="True" OnSelectedIndexChanged="DdlUnidadeNegocio_SelectedIndexChanged"
                        Width="597px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlEmpresa" runat="server" Width="597px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Conta:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtConta" runat="server" Width="100px" CssClass="texto" Enabled="False" />
                    <asp:TextBox ID="txtNomeConta" runat="server" Width="436px" CssClass="texto" Enabled="False" />
                    <asp:ImageButton ID="ImgConta" OnClick="ImgConta_Click" runat="server" ImageAlign="AbsMiddle"
                        ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png" data-ToolTip="default" ToolTip="Conta para lançamento contábil." />
                    <asp:ImageButton ID="ImgLimparConta" OnClick="ImgLimparConta_Click" runat="server"
                        data-ToolTip="default" ToolTip="Limpar" ImageAlign="AbsMiddle" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_limpar.png" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Depreciação de Débito:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtContaDepreciacaoDebito" runat="server" Enabled="False" Width="100px"
                        CssClass="texto" data-ToolTip="default" />
                    <asp:TextBox ID="txtNomeDepreciacaoDebito" runat="server" Enabled="False" Width="436px"
                        CssClass="texto" />
                    <asp:ImageButton ID="ImgDepDebito" OnClick="ImgDepDebito_Click" runat="server" ImageAlign="AbsMiddle"
                        ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png" data-ToolTip="default" ToolTip="Conta contábil quando a depreciação for a débito." />
                    <asp:ImageButton ID="ImgLimparDepDebito" OnClick="ImgLimparDepDebito_Click" runat="server"
                        data-ToolTip="default" ToolTip="Limpar" ImageAlign="AbsMiddle" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_limpar.png" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Depreciação de Crédito:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtContaDepreciacaoCredito" runat="server" Enabled="False" Width="100px"
                        CssClass="texto" />
                    <asp:TextBox ID="txtNomeDepreciacaoCredito" runat="server" Enabled="False" Width="436px"
                        CssClass="texto" />
                    <asp:ImageButton ID="ImgDepCredito" OnClick="ImgDepCredito_Click" runat="server"
                        ImageAlign="AbsMiddle" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png" data-ToolTip="default" ToolTip="Conta contábil quando a depreciação for a crédito." />
                    <asp:ImageButton ID="ImgLimparDepCredito" OnClick="ImgLimparDepCredito_Click" runat="server"
                        data-ToolTip="default" ToolTip="Limpar" ImageAlign="AbsMiddle" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_limpar.png" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Conta de Baixa:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtContaDeBaixa" runat="server" Enabled="False" Width="100px" CssClass="texto" />
                    <asp:TextBox ID="txtNomeContaDeBaixa" runat="server" Enabled="False" Width="436px"
                        CssClass="texto" />
                    <asp:ImageButton ID="imgContaDeBaixa" OnClick="imgContaDeBaixa_Click" runat="server"
                        ImageAlign="AbsMiddle" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png" data-ToolTip="default" ToolTip="Conta contábil para registrar a baixa do bem." />
                    <asp:ImageButton ID="imgLimparContaDeBaixa" OnClick="imgLimparContaDeBaixa_Click"
                        data-ToolTip="default" ToolTip="Limpar" runat="server" ImageAlign="AbsMiddle"
                        ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_limpar.png" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Conta de Transferência:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtContaTransferencia" runat="server" Enabled="False" Width="100px"
                        CssClass="texto" />
                    <asp:TextBox ID="txtNomeContaTransferencia" runat="server" Enabled="False" Width="436px"
                        CssClass="texto" />
                    <asp:ImageButton ID="imgContaTransferencia" OnClick="imgContaTransferencia_Click"
                        runat="server" ImageAlign="AbsMiddle" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png" data-ToolTip="default" ToolTip="Conta contábil para registrar a transferência do bem." />
                    <asp:ImageButton ID="imgLimparContaTransferencia" OnClick="imgLimparContaTransferencia_Click"
                        data-ToolTip="default" ToolTip="Limpar" runat="server" ImageAlign="AbsMiddle"
                        ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_limpar.png" />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="GridAtivosXContas" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="GridAtivosXContas_SelectedIndexChanged">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:CommandField SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button"></asp:CommandField>
                        <asp:BoundField DataField="Conta_Id" HeaderText="Conta">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Titulo" HeaderText="Nome">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="DepreciacaoDebito" HeaderText="Deprec. D&#233;b.">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="DepreciacaoCredito" HeaderText="Deprec. Cr&#233;d.">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="ContaDeBaixa" HeaderText="Baixa.">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="ContaDeTransferencia" HeaderText="Transferência">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaPlanoDeContas ID="ucConsultaPlanoDeContas" runat="server" />
</asp:Content>
