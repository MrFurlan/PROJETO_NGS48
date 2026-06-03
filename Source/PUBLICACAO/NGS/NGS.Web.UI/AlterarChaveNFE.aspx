<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="AlterarChaveNFE.aspx.vb" Inherits="NGS.Web.UI.AlterarChaveNFE" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <script type="text/javascript">
        function pageLoadNota() {
            $("#txtChaveNFe").live('keypress', function (e) {
                var code = (e.keyCode ? e.keyCode : e.which);
                if (code == 13) {
                    $("#MainContent_imgVerificarChaveNFE").click();
                    return false;
                }
            });

            $("#fupArquivo").focus(function () {
                return false;
            });
        }

        $(document).ready(function () {
            pageLoadNota();
            var prmNota = Sys.WebForms.PageRequestManager.getInstance();
            prmNota.add_endRequest(pageLoadNota);
        });
    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmAlterarChaveNFE" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlAlterarChaveNFE" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Alterar Chave NFE/CTE
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li id="Li1" runat="server">
                            <asp:LinkButton class="iconAtualizar" ID="lnkAtualizar" runat="server" Text="Atualizar" />
                        </li>
                        <li id="Li2" runat="server">
                            <asp:LinkButton class="iconConsultar" ID="lnkConsultar" runat="server" Text="Consultar" />
                        </li>
                        <li id="Li3" runat="server">
                            <asp:LinkButton class="iconLimpar" ID="lnkLimpar" runat="server" Text="Limpar" />
                        </li>
                        <li id="Li4" runat="server">
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
            <div class="row" id="divNFE" runat="server" visible="False">
                <div class="borda">
                    <div class="row">
                        <strong><span style="font-size: 6pt;">CHAVE DE ACESSO DA NF-e CONSULTA NO SITE: WWW.NFE.FAZENDA.GOV.BR</span></strong>
                    </div>
                    <div class="row">
                        <asp:TextBox ID="txtChaveNFe" CssClass="chaveNFe" ClientIDMode="Static" runat="server"
                            Width="320px" ReadOnly="True" TabIndex="2" />
                        <asp:LinkButton ID="lnkVerificarChaveNFE" CssClass="lnk"
                            data-tooltip="default" ToolTip="Consultar/Validar Chave Eletrônica" runat="server" Text=" &gt; "
                            OnClick="lnkVerificarChaveNFE_Click">
                                    <i class="fa fa-arrow-right seta"></i>
                        </asp:LinkButton>
                    </div>
                    <div class="row" id="divArquivo" runat="server" visible="True">
                        <div class="collbl">
                            Arquivo:
                        </div>
                        <div class="coltxt">
                            <uc:File ID="ucFile" runat="server" />
                        </div>
                    </div>
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaPedidosxNotas ID="ucConsultaPedidosXNotas" runat="server" />
</asp:Content>
