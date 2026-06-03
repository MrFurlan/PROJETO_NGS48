<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master" CodeBehind="RecusaNotaFiscal.aspx.vb" Inherits="NGS.Web.UI.RecusaNotaFiscal" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngRecusaNotaFiscal" runat="server" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlRecusaNotaFiscal" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Recusa de NFe/CTe
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconAtualizar" ID="lnkAtualizar" runat="server" Text="Enviar" />
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
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="634px" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Chave Eletrônica:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtChaveNFe" ClientIDMode="Static" runat="server" CssClass="chaveNFe"
                        Width="400px" data-ToolTip="default" ToolTip="Código de identificação da NFe ou CTe perante a Receita Federal." />
                    <asp:LinkButton ID="lnkVerificarChaveNFE" CssClass="lnk" 
                        data-tooltip="default" ToolTip="Consultar/Validar Chave Eletrônica." runat="server" Visible="false" Text=" &gt; "
                        OnClick="lnkVerificarChaveNFE_Click">
                        <i class="fa fa-arrow-right seta"></i>
                    </asp:LinkButton>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Tipo de Operação:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdNFe" runat="server" Text="NFe - Operação não Realizada" GroupName="Ope" data-ToolTip="default" ToolTip="Informar o tipo de Operação." AutoPostBack="True" />
                    <asp:RadioButton ID="rdCTe" runat="server" Text="CTe - Prestação de Serviço em Desacordo" GroupName="Ope" data-ToolTip="default" ToolTip="Informar o tipo de Operação." AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Justificativa:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtjustificativa" runat="server" Width="400px" TextMode="MultiLine" data-ToolTip="default" ToolTip="Informe a justificativa para a operação." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
