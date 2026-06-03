<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="GeracaoBradesco.aspx.vb" Inherits="NGS.Web.UI.GeracaoBradesco" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngGeracaoBradesco" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlGeracaoBradesco" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Gera Arquivo para o Banco Bradesco
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconRelatorio" ID="lnkRelatorio" Text="Relatório" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconConsultar" ID="lnkGerar" Text="Gerar Arquivo" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconRelatorio" ID="lnkBaixa" Text="Baixa por Lote" runat="server" />
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
                    <asp:DropDownList ID="ddlUnidadeGeracaoBradesco" runat="server" AutoPostBack="True"
                        OnSelectedIndexChanged="ddlUnidadeGeracaoBradesco_SelectedIndexChanged" TabIndex="10"
                        Width="594px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlEmpresaConsultaTitulos" runat="server" AutoPostBack="True"
                        OnSelectedIndexChanged="DdlEmpresaConsultaTitulos_SelectedIndexChanged" TabIndex="2"
                        Width="594px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Conta:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlContaConsulta" runat="server" TabIndex="2" Width="594px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Periodo Inicial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPeriodoInicialConsultaTitulos" runat="server" Width="116px" CssClass="calendario"
                        TabIndex="5" data-ToolTip="default" ToolTip="Periodo Inicial" />
                </div>
                <div class="collbl">
                    Periodo Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPeriodoFinalConsultaTitulos" runat="server" CausesValidation="True"
                        CssClass="calendario" TabIndex="6" Width="116px" data-ToolTip="default" ToolTip="Informar o período final de consulta." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Tipo/relatorio:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="RbGeral" runat="server" Checked="True" GroupName="TipoRel" OnCheckedChanged="RbGeral_CheckedChanged"
                        Text="Geral" data-ToolTip="default" ToolTip="Selecionar o tipo de documento para gerar o relatório." />
                    <asp:RadioButton ID="RbEnviado" runat="server" GroupName="TipoRel" OnCheckedChanged="RbEnviado_CheckedChanged"
                        Text="Enviado" data-ToolTip="default" ToolTip="Selecionar o tipo de documento para gerar o relatório." />
                    <asp:RadioButton ID="RbBaixado" runat="server" GroupName="TipoRel" OnCheckedChanged="RbBaixado_CheckedChanged"
                        Text="Baixado" data-ToolTip="default" ToolTip="Selecionar o tipo de documento para gerar o relatório." />
                    <asp:RadioButton ID="RbAEnviar" runat="server" GroupName="TipoRel" Text="A Enviar" data-ToolTip="default"
                        ToolTip="Selecionar o tipo de documento para gerar o relatório." />
                    <asp:RadioButton ID="Rb_NaoProc" runat="server" GroupName="TipoRel" OnCheckedChanged="Rb_NaoProc_CheckedChanged"
                        Text="Não Processado" data-ToolTip="default" ToolTip="Selecionar o tipo de documento para gerar o relatório." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
