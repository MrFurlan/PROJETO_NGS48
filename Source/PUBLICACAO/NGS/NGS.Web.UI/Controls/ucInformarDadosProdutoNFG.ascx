<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucInformarDadosProdutoNFG.ascx.vb"
    Inherits="NGS.Web.UI.ucInformarDadosProdutoNFG" %>

<div id="divInformarDadosProdutoNFG" class="uc" title="Informe o produto" style="display: none;">
    <asp:UpdatePanel ID="updpnlProdutoNFG" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <asp:HiddenField ID="PosicaoItem" runat="server" />
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconNovo" runat="server">
                            <asp:LinkButton ID="lnkGravar" Text="Adicionar" runat="server" />
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li class="iconFechar" runat="server">
                            <asp:LinkButton ID="lnkFechar" Text="Fechar" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Grupo:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbGrupoProduto" runat="server" AutoPostBack="True" OnSelectedIndexChanged="cmbGrupoProduto_SelectedIndexChanged"
                        TabIndex="13" Width="620px" />
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Produto:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbProdutos" runat="server" Width="590px" OnSelectedIndexChanged="cmbProdutos_SelectedIndexChanged"
                        AutoPostBack="True" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="BtnConsultaProduto" CssClass="btn" runat="server" UseSubmitBehavior="False"
                        Text=" > " data-ToolTip="default" ToolTip="Nome do produto." />
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Operação:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtOperacao" runat="server" Font-Names="monospace" Enabled="false" Width="589px"
                        data-ToolTip="default" ToolTip="Referente a operação de criação da nota." />
                </div>
                <div class="coltxt">
                    <asp:ImageButton ID="imbOperacao" runat="server" Style="cursor: pointer; border: 0 none;"
                        data-ToolTip="default" ToolTip="Consulta de Operações" ImageAlign="AbsMiddle"
                        ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png" OnClick="imbOperacao_Click" />
                </div>
            </div>
            <div class="row" id="divProdutoCusto" runat="server" visible="false">
                <div class="collbluc">
                    Grupo:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbGrupoProdutoCusto" runat="server" AutoPostBack="True" OnSelectedIndexChanged="cmbGrupoProdutoCusto_SelectedIndexChanged" Width="620px" />
                </div>
                <div class="collbluc">
                    Produto:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbProdutosCusto" runat="server" Width="590px" OnSelectedIndexChanged="cmbProdutosCusto_SelectedIndexChanged" AutoPostBack="True" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="BtnConsultaProdutoCusto" CssClass="btn" runat="server" UseSubmitBehavior="False"
                        Text=" > " data-ToolTip="default" ToolTip="Nome do produto." />
                </div>
            </div>
            <div id="divCentroDeCusto" class="row" runat="server">
                <div class="collbluc">
                    Centro de Custo:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbCentroCusto" runat="server" Font-Names="monospace" Width="620px" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
