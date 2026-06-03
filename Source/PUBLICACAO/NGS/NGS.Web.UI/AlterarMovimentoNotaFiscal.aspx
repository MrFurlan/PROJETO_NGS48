<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="AlterarMovimentoNotaFiscal.aspx.vb" Inherits="NGS.Web.UI.AlterarMovimentoNotaFiscal" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmAlterarMovimentoNotaFiscal" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlAlterarMovimentoNotaFiscal" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Alterar Movimento Nota Fiscal
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconAtualizar" ID="lnkAtualizar" runat="server" Text="Atualizar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAtualizar" ID="lnkAjustarHora" runat="server" Text="Ajustar Hora" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconConsultar" ID="lnkConsultar" runat="server" Text="Consultar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconLimpar" ID="lnkLimpar" runat="server" Text="Limpar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" runat="server" Text="Ajuda" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUnidadeDeNegocio" runat="server" Width="633px" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="633px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                    <asp:TextBox ID="txtCliente" runat="server" Width="593px" Enabled="False" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnCliente" runat="server" Text=">" CssClass="btn" UseSubmitBehavior="False"
                        data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    UF:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtUf" runat="server" Width="100px" MaxLength="2" data-ToolTip="default"
                        ToolTip="Abreviatura do estado." />
                </div>
                <div class="collbl">
                    Data Inicial:
                </div>
                <div class="coltxt" style="width: 110px;">
                    <asp:TextBox ID="txtDataInicial" runat="server" CssClass="calendario" Width="79px"
                        data-ToolTip="default" ToolTip="Período inicial da pesquisa." />
                </div>
                <div class="collbl">
                    Data Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" runat="server" CssClass="calendario" Width="100px"
                        data-ToolTip="default" ToolTip="Período final da pesquisa." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    E/S:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtES" runat="server" Width="100px" MaxLength="1" data-ToolTip="default"
                        ToolTip="Informar se a nota é de entrada ou saída." />
                </div>
                <div class="collbl">
                    Nota:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNota" Style="text-align: right;" runat="server" class="txtNumerico"
                        Width="100px" data-ToolTip="default" ToolTip="Número da Nota Fiscal." />
                </div>
                <div class="collbl">
                    Série:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtSerie" runat="server" Width="100px" MaxLength="3" data-ToolTip="default"
                        ToolTip="Série da Nota Fiscal." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Situação:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbSituacao" runat="server" Width="240px" Enabled="False" />
                </div>
            </div>
            <div class="row" id="divDatas" runat="server" visible="False">
                <div class="collbl">
                    Data Emissão:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataEmissao" runat="server" CssClass="calendario" Width="79px" />
                </div>
                <div class="collbl">
                    Data E/S:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataES" runat="server" CssClass="calendario" Width="79px" />
                </div>
                <div class="collbl">
                    Data e Hora Inclusão:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInclusao" runat="server" Width="100px" Enabled="False" /> 
                    <asp:ImageButton ID="imgMenosHora" runat="server" Width="16px" Height="16px" ImageUrl="~/Images/ico-menos.gif" OnClick="imgMenosHora_Click" data-ToolTip="default" ToolTip="Diminuir 1(uma) hora." />
                    <asp:ImageButton ID="imgMaisHora" runat="server" Width="16px" Height="16px" ImageUrl="~/Images/ico-mais.gif" OnClick="imgMaisHora_Click" data-ToolTip="default" ToolTip="Aumentar 1(uma) hora." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaPedidosxNotas ID="ucConsultaPedidosXNotas" runat="server" />
</asp:Content>
