<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="RelatorioNotasXNotas.aspx.vb" Inherits="NGS.Web.UI.RelatorioNotasXNotas" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngRelatorioNotasXNotas" runat="server" AsyncPostBackTimeout="900" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlRelatorioNotasXNotas" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Relatório de Notas X Notas
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconExcel" ID="lnkRelatório" runat="server"
                                Text="Relatório" />
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" runat="server"
                                Text="Ajuda" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUnidade" runat="server" Width="634px" OnSelectedIndexChanged="DdlUnidade_SelectedIndexChanged"
                        AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkUnificarEmpresa" runat="server" Text="Empresa  :" ToolTip="Unificar os CNPJs da empresa." />
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlEmpresa" runat="server" Width="634px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                    <asp:TextBox ID="txtCliente" runat="server" Width="600px" Enabled="false" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="cmdCliente" CssClass="btn" OnClick="cmdCliente_Click" runat="server"
                        Text=">" data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Operação:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbOperacao" runat="server" Width="294px" AutoPostBack="True"
                        OnSelectedIndexChanged="cmbOperacao_SelectedIndexChanged" />
                    <asp:DropDownList ID="cmbSubOperacao" runat="server" Width="295px" />
                </div>
            </div>
            <div class="row">
                <div class="coltxt lg">
                    <uc:SelecaoProduto ID="ucSelecaoProduto" runat="server" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Notas:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="RadEntradas" runat="server" Text="Entradas" Checked="True" GroupName="Notas" data-ToolTip="default" ToolTip="Informar se a nota é de entrada ou saída." />
                    <asp:RadioButton ID="RadSaidas" runat="server" Text="Saidas" GroupName="Notas" data-ToolTip="default" ToolTip="Informar se a nota é de entrada ou saída." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Período Inicial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicial" CssClass="calendario" runat="server" Width="96px" data-ToolTip="default" ToolTip="Informar o data inicial de consulta." />
                </div>
                <div class="collbl">
                    Período Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" CssClass="calendario" runat="server" Width="96px" data-ToolTip="default" ToolTip="Informar o data final de consulta." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkNossaEmissao" Text="Nossa Emissão:" runat="server" AutoPostBack="True" data-ToolTip="default" ToolTip="Apenas Notas Fiscais Nossa Emissão." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkAllTipos" data-ToolTip="default" ToolTip="Seleciona todos os Tipos de Documentos."
                        Text="Tipo Documento:" runat="server" AutoPostBack="True" OnCheckedChanged="chkAllTipos_CheckedChanged" />
                </div>
                <div class="coltxt" style="line-height: 12px;">
                    <asp:CheckBoxList ID="chkTipoDeDocumento" runat="server" RepeatColumns="3" data-ToolTip="default"
                        ToolTip="Selecionar qual ou quais os tipos de documentos." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
