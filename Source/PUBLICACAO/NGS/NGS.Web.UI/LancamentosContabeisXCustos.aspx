<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="LancamentosContabeisXCustos.aspx.vb" Inherits="NGS.Web.UI.LancamentosContabeisXCustos" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngLancamentosContabeisXCustos" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlLancamentosContabeisXCustos" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Lançamentos Contábeis X Custos
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
                    <asp:DropDownList ID="ddlUnidade" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlUnidade_SelectedIndexChanged"
                        TabIndex="2" Width="726px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlEmpresa_SelectedIndexChanged"
                        TabIndex="3" Width="726px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Movimento:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtMovimento" runat="server" AutoPostBack="True" CssClass="calendario"
                        OnTextChanged="txtMovimento_TextChanged" Width="88px" data-ToolTip="default"
                        ToolTip="Data do lançamento." />
                </div>
                <div class="collbl">
                    Lote:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlLote" runat="server" AutoPostBack="true" OnSelectedIndexChanged="ddlLote_SelectedIndexChanged"
                        TabIndex="1" Width="110px" />
                </div>
                <div class="collbl">
                    Seq:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlSequencia" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlSequencia_SelectedIndexChanged"
                        Width="110px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    [D] Conta:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtContaDebito" runat="server" Enabled="False" Width="693px" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnContasDebito" runat="server" Text=">" CssClass="btn" data-ToolTip="default"
                        ToolTip="Conta Contábil a débito." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    [E] Cliente:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtClienteDebito" runat="server" Enabled="False" Width="693px" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnClientesDebito" runat="server" Text=">" CssClass="btn" data-ToolTip="default"
                        ToolTip="Identificação do cliente à débito." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    [B]C.Custo:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlCustoDebito" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlCustoDebito_SelectedIndexChanged"
                        TabIndex="6" Width="726px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    [C] Conta:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtContaCredito" runat="server" Enabled="False" Width="693px" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnContasCredito" runat="server" Text=">" CssClass="btn" data-ToolTip="default"
                        ToolTip="Conta Contábil a credito." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    [R] Cliente:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtClienteCredito" runat="server" Enabled="False" Width="693px" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnClientesCredito" runat="server" Text=">" CssClass="btn" data-ToolTip="default"
                        ToolTip="Identificação do cliente à credito." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    [E] C.Custo:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlCustoCredito" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlCustoCredito_SelectedIndexChanged"
                        TabIndex="11" Width="726px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Grupo:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlGrupo" runat="server" AutoPostBack="True" Width="726px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Produto:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlProduto" runat="server" TabIndex="12" Width="726px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Valor R$:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtValorDocumento" runat="server" />
                    <asp:TextBox ID="txtValor" runat="server" CssClass="txtDecimal" data-ToolTip="default"
                        ToolTip="Valor referente ao lançamento." />
                </div>
                <div class="collbl">
                    Quantidade:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtQuantidade" runat="server" CssClass="txtDecimal4" Style="color: blue; text-align: right"
                        Width="118px" data-ToolTip="default" ToolTip="Quantidade de produto." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Histórico:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlHistorico" runat="server" Width="726px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Totalizadores:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtTotalDebitos" runat="server" Font-Bold="True" placeholder="Débitos"
                        Enabled="False" TabIndex="100" Width="100px" data-ToolTip="default" ToolTip="Valores a débito, crédito e diferenças." />
                    <asp:TextBox ID="txtTotalCreditos" runat="server" Font-Bold="True" placeholder="Créditos"
                        Enabled="False" TabIndex="101" Width="100px" data-ToolTip="default" ToolTip="Valores a débito, crédito e diferenças." />
                    <asp:TextBox ID="txtDiferenca" runat="server" Font-Bold="True" placeholder="Diferença"
                        Enabled="False" TabIndex="102" Width="100px" data-ToolTip="default" ToolTip="Valores a débito, crédito e diferenças." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Complemento:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtComplemento" runat="server" Rows="2" TabIndex="15" TextMode="MultiLine"
                        Width="723px" data-ToolTip="default" ToolTip="Preencher quando houver informações relevantes." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaPlanoDeContas ID="ucConsultaPlanoDeContas" runat="server" />
</asp:Content>
