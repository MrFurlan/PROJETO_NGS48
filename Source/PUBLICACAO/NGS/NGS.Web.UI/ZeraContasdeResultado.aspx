<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="ZeraContasdeResultado.aspx.vb" Inherits="NGS.Web.UI.ZeraContasdeResultado" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
        <style type="text/css">
        #meioconteudo {
            width: 712px !important;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmZeraContasdeResultado" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlZeraContasdeResultado" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Zeramento das Contas de Resultado
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconLimpar" ID="lnkZerarContas" runat="server" Text="Zerar Contas"
                                OnClick="lnkZerarContas_Click" UseSubmitBehavior="False" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconExcluir" ID="lnkExcluir" runat="server" Text="Excluir"
                                OnClientClick="return msgconfirm(this);" />
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
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="570px" />
                </div>
            </div>
            <div class="row">
                <div class="coltxt" style="margin-left: 126px;">
                    <asp:CheckBox ID="chkContasDeEstoque" runat="server" AutoPostBack="True" OnCheckedChanged="chkContasDeEstoque_CheckedChanged"
                        Text="Contas De Estoque" data-ToolTip="default" ToolTip="Selecionar somente contas de estoque." />
                </div>
                <div class="coltxt" style="margin-left: 126px;">
                    <asp:Label ID="lblLote" runat="server" Font-Bold="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Conta Resultado:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtConta0" runat="server" OnTextChanged="txtConta0_TextChanged"
                        data-ToolTip="default" ToolTip="Informar a conta de resultado para zeramento." />
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtConta1" runat="server" Width="22px" data-tooltip="default"
                        ToolTip="Informar a conta de resultado para zeramento." />
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtConta2" runat="server" Width="401px" ReadOnly="True" data-ToolTip="default"
                        ToolTip="Informar a conta de resultado para zeramento." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Período:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicio" runat="server" CssClass="calendario" data-ToolTip="default"
                        ToolTip="Data inicial a final para consulta." />
                    <asp:TextBox ID="txtDataFim" runat="server" CssClass="calendario" data-ToolTip="default"
                        ToolTip="Data inicial a final para consulta." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
