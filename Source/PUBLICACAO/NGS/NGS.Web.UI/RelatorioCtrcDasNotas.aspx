<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="RelatorioCtrcDasNotas.aspx.vb" Inherits="NGS.Web.UI.RelatorioCtrcDasNotas" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngRelatorioCtrcDasNotas" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlRelatorioCtrcDasNotas" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Relatório CTRC das Notas
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconRelatorio" runat="server" style="width: 10%;">
                            <asp:LinkButton ID="lnkRelatorio" Text="Gerar PDF" runat="server" />
                        </li>
                        <li class="iconRelatorio" runat="server" style="width: 12%;">
                            <asp:LinkButton ID="lnkExcel" Text="Gerar EXCEL" runat="server" />
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li class="iconAjuda" runat="server">
                            <asp:LinkButton ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkUnificarEmpresa" runat="server" Text="Empresa  :" ToolTip="Unificar os CNPJs da empresa." />
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="598px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCliente" runat="server" Width="558px" Enabled="False" />
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnConsultaCliente" CssClass="btn" runat="server" Text=">" UseSubmitBehavior="False"
                        data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Transportador:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoTransp" runat="server" />
                    <asp:TextBox ID="txtTransportador" runat="server" Width="558px" Enabled="False" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnConsultaTransp" CssClass="btn" runat="server" Text=">" UseSubmitBehavior="False"
                        data-ToolTip="default" ToolTip="Nome da pessoa/empresa responsável pelo transporte da mercadoria." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Nota:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNota" runat="server" Width="90px" data-ToolTip="default" ToolTip="Inserir o número da NF." />
                </div>
                <div class="collbl">
                    Pedido:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPedido" runat="server" Width="80px" data-ToolTip="default" ToolTip="Informar o número do pedido." />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnPedido" CssClass="btn" runat="server" Text=">" UseSubmitBehavior="False"
                        data-ToolTip="default" ToolTip="Informar o número do pedido." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtData1" CssClass="calendario" runat="server" Width="70px" data-ToolTip="default"
                        ToolTip="Informar a data inicial e fnal da consulta." />
                    &nbsp;&nbsp;á&nbsp;
                    <asp:TextBox ID="txtData2" CssClass="calendario" runat="server" Width="70px" data-ToolTip="default"
                        ToolTip="Informar a data inicial e fnal da consulta." />
                </div>
            </div>
            <div class="row">
                <div class="coltxt lg">
                    <uc:SelecaoProduto ID="ucSelecaoProduto" runat="server" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    PF/PJ:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdoTodos" runat="server" Text="Todos" GroupName="grpFisicaJuridica"
                        Checked="true" data-ToolTip="default" ToolTip="Informar se o frete é por pessoa Física ou Jurídica." />
                    <asp:RadioButton ID="rdoFisica" runat="server" Text="Física" GroupName="grpFisicaJuridica"
                        data-ToolTip="default" ToolTip="Informar se o frete é por pessoa Física ou Jurídica." />
                    <asp:RadioButton ID="rdoJuridica" runat="server" Text="Jurídica" GroupName="grpFisicaJuridica"
                        data-ToolTip="default" ToolTip="Informar se o frete é por pessoa Física ou Jurídica." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Formato:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdoNotaXCtrc" runat="server" Text="Nota Fiscal X Conhecimento"
                        GroupName="grpFormato" Checked="true" data-ToolTip="default" ToolTip="Informar o formato do relatório." />
                    <asp:RadioButton ID="rdoCircXComp" runat="server" Text="Circulação X Comprovação"
                        GroupName="grpFormato" data-ToolTip="default" ToolTip="Informar o formato do relatório." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    E/S:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdoEntrada" runat="server" Text="Entrada" GroupName="grpEntradaSaida"
                        Checked="true" data-ToolTip="default" ToolTip="Informar se a nota é de entrada ou saída." />
                    <asp:RadioButton ID="rdoSaida" runat="server" Text="Saída" GroupName="grpEntradaSaida"
                        data-ToolTip="default" ToolTip="Informar se a nota é de entrada ou saída." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaPedidos ID="ucConsultaPedidos" runat="server" />
</asp:Content>
