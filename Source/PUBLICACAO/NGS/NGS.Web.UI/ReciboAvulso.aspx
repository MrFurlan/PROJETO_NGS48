<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="ReciboAvulso.aspx.vb" Inherits="NGS.Web.UI.ReciboAvulso" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngReciboAvulso" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlReciboAvulso" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Recibo Avulso
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconRelatorio" runat="server">
                            <asp:LinkButton ID="lnkEmitir" runat="server" Text="Emitir" />
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" runat="server" Text="Limpar" />
                        </li>
                        <li class="iconAjuda" runat="server">
                            <asp:LinkButton ID="lnkAjuda" runat="server" Text="Ajuda" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUnidadeNegocio" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlUnidadeNegocio_SelectedIndexChanged"
                        Width="589px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" AutoPostBack="True" Width="589px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtClientes" runat="server" Width="549px" data-ToolTip="default"
                        ToolTip="Selecionar o cliente desejado." />
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="ddlBuscaCliente" OnClick="ddlBuscaCliente_Click" runat="server" Text=">"
                        CausesValidation="False" UseSubmitBehavior="False" CssClass="btn" data-ToolTip="default"
                        ToolTip="Selecionar o cliente desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Endereço:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="TxtEndereco" runat="server" Width="580px" data-ToolTip="default"
                        ToolTip="Local que a empresa está situada." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    CPF/CNPJ:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="TxtCpfCnpj" runat="server" Width="191px" data-ToolTip="default"
                        ToolTip="Número do cadastro de pessoa física ou jurídica perante a Receita Federal." />
                </div>
                <div class="collbl">
                    Numero:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="TxtNumero" CssClass="txtNumerico" runat="server" Width="85px" data-ToolTip="default" ToolTip="Identificação do imóvel em seu endereço." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Bairro:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="TxtBairro" runat="server" Width="191px" data-ToolTip="default" ToolTip="Referente ao local em que o endereço está situado na cidade." />
                </div>
                <div class="collbl">
                    CEP:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="TxtCEP" CssClass="txtCep" runat="server" Width="85px" data-ToolTip="default"
                        ToolTip="Código de endereçamento postal." />
                </div>
                <div class="coltxt">
                    <asp:ImageButton ID="imgCep" runat="server" ImageAlign="AbsMiddle" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png"
                        OnClick="imgCep_Click" data-ToolTip="default" ToolTip="Procura por endereço com base no Cep informado"
                        CssClass="imgconsultar" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cidade:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="TxtCidade" runat="server" Width="191px" data-ToolTip="default" ToolTip="Localização da empresa dentro do Estado." />
                </div>
                <div class="collbl">
                    UF:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="TxtEstado" runat="server" Width="85px" data-ToolTip="default" ToolTip="Abreviatura do estado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Complemento:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="TxtComplemento" runat="server" Width="191px" data-ToolTip="default"
                        ToolTip="Complemento do endereço." />
                </div>
                <div class="collbl">
                    Telefone:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="TxtTelefone" CssClass="txtFone" runat="server" Width="85px" data-ToolTip="default"
                        ToolTip="Contato da empresa." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Número do Recibo:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNumeroRecibo" CssClass="txtNumerico" runat="server" Width="191px" data-ToolTip="default" ToolTip="Numeração do Recibo." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Total do Recibo:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtTotalRecibo" CssClass="txtDecimal" runat="server" Width="191px"
                        AutoPostBack="True" data-ToolTip="default" ToolTip="Valor total do recibo." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Vencimento:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtData" runat="server" CssClass="calendario" Width="85px" data-ToolTip="default"
                        ToolTip="Data de vencimento do pagamento." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Referente:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="TxtReferente" runat="server" Width="576px" data-ToolTip="default" TextMode="MultiLine" ToolTip="" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
