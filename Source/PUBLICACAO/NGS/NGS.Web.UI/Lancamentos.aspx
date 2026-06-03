<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="Lancamentos.aspx.vb" Inherits="NGS.Web.UI.Lancamentos" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <script type="text/javascript">
        function pageLoadLancamentos() {
            $("#txtMovimento").datepicker({
                dateFormat: 'dd/mm/yy',
                dayNames: ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado', 'Domingo'],
                dayNamesMin: ['D', 'S', 'T', 'Q', 'Q', 'S', 'S', 'D'],
                dayNamesShort: ['Dom', 'Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb', 'Dom'],
                monthNames: ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'],
                monthNamesShort: ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'],
                nextText: 'Próximo',
                prevText: 'Anterior',
                showOn: "button",
                buttonImage: "Images/calendar.png",
                buttonImageOnly: true
            });

            $("#txtMovimento").setMask("date");
        }

        $(document).ready(function () {
            pageLoadLancamentos();
            var prmLancamentos = Sys.WebForms.PageRequestManager.getInstance();
            prmLancamentos.add_endRequest(pageLoadLancamentos);
        });
    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngLancamentos" runat="server" AsyncPostBackTimeout="10000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlLancamentos" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Lançamentos Contábeis
                <div style="float: right; padding-right: 10px;">
                    <asp:Image ID="imgUsuarioIncl" runat="server" Height="20px" ImageAlign="AbsMiddle"
                        ImageUrl="~/Images/man2.png" Width="18px" />
                    <asp:Label ID="lblUsuario" runat="server" ForeColor="Yellow" Font-Bold="True" Font-Size="11px" />
                </div>
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
                        <li runat="server" style="width: 9%;">
                            <asp:LinkButton class="iconLimpar" ID="lnkLimpar" runat="server" Text="Limpar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="painelleft" style="width: 83%;">
                <div class="row">
                    <div class="collbl">
                        Filtro:
                    </div>
                    <div class="coltxt">
                        <asp:CheckBox ID="ChkAproveitarDados" runat="server" Text="Aproveitar dados para novo lançamento." Style="float: left;"
                            data-ToolTip="default" ToolTip="Com esta opção é possivel manter os dados para o próximo lançamento." />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Unidade:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlUnidade" runat="server" Width="600px" OnSelectedIndexChanged="ddlUnidade_SelectedIndexChanged"
                            AutoPostBack="True" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Empresa:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlEmpresa" runat="server" Width="600px" OnSelectedIndexChanged="ddlEmpresa_SelectedIndexChanged"
                            AutoPostBack="True" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        [D] Conta:
                    </div>
                    <div class="coltxt">
                        <asp:HiddenField ID="txtTemClienteDebito" runat="server" />
                        <asp:HiddenField ID="txtCodigoContaDebito" runat="server" />
                        <asp:TextBox ID="txtContaDebito" runat="server" Width="560px" Enabled="False" />
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
                        <asp:HiddenField ID="txtCodigoClienteDebito" runat="server" />
                        <asp:TextBox ID="txtClienteDebito" runat="server" Width="560px" Enabled="False" />
                    </div>
                    <div class="coltxt">
                        <asp:Button ID="btnClientesDebito" runat="server" Text=">" CssClass="btn" data-ToolTip="default"
                            ToolTip="Identificação do cliente à débito." />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        [B] Centro De Custo:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlCustoDebito" runat="server" Width="600px" OnSelectedIndexChanged="ddlCustoDebito_SelectedIndexChanged"
                            AutoPostBack="True" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        [C] Conta:
                    </div>
                    <div class="coltxt">
                        <asp:HiddenField ID="txtTemClienteCredito" runat="server" />
                        <asp:HiddenField ID="txtCodigoContaCredito" runat="server" />
                        <asp:TextBox ID="txtContaCredito" runat="server" Width="560px" Enabled="False" />
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
                        <asp:HiddenField ID="txtCodigoClienteCredito" runat="server" />
                        <asp:TextBox ID="txtClienteCredito" runat="server" Width="560px" Enabled="False" />
                    </div>
                    <div class="coltxt">
                        <asp:Button ID="btnClientesCredito" runat="server" Text=">" CssClass="btn" data-ToolTip="default"
                            ToolTip="Identificação do cliente à credito." />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        [E] Centro De Custo:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlCustoCredito" runat="server" Width="600px" OnSelectedIndexChanged="ddlCustoCredito_SelectedIndexChanged"
                            AutoPostBack="True" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Grupo:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="DdlGrupo" runat="server" Width="600px" AutoPostBack="True" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Produto:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlProduto" runat="server" Width="600px" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Pedido:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtPedido" runat="server" Enabled="False" CssClass="txtInteiro" />
                        <asp:Button ID="btnPedido" runat="server" OnClick="btnPedido_Click" Text=">" CssClass="btn"
                            data-ToolTip="default" ToolTip="Número do pedido a ser consultado." />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Valor R$:
                    </div>
                    <div class="coltxt">
                        <asp:HiddenField ID="txtValorDocumento" runat="server" />
                        <asp:TextBox ID="txtValor" CssClass="txtDecimal" runat="server" data-ToolTip="default"
                            ToolTip="Valor referente ao lançamento." />
                    </div>
                    <div class="collbl">
                        Quantidade:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtQuantidade" CssClass="txtDecimal4" runat="server" Style="color: blue;
                            text-align: right" Width="118px" data-ToolTip="default" ToolTip="Quantidade de produto." />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Histórico:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlHistorico" runat="server" Width="600px" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Complemento:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtComplemento" runat="server" Width="595px" TextMode="MultiLine"
                            Rows="2" data-ToolTip="default" ToolTip="Preencher quando houver informações relevantes." />
                    </div>
                </div>
            </div>
            <div class="painelright" style="width: 16%;">
                <div class="subtitulodiv">
                    Movimento:
                </div>
                <div class="row">
                    <div class="coltxt">
                        <asp:TextBox ID="txtMovimento" runat="server" Width="88px" AutoPostBack="True" OnTextChanged="txtMovimento_TextChanged" ClientIDMode="Static" 
                            data-ToolTip="default" ToolTip="" />
                    </div>
                </div>
                <div class="subtitulodiv">
                    Lote:
                </div>
                <div class="row">
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlLote" runat="server" Width="110px" OnSelectedIndexChanged="ddlLote_SelectedIndexChanged"
                            AutoPostBack="true" />
                    </div>
                </div>
                <div class="subtitulodiv">
                    Sequência:
                </div>
                <div class="row">
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlSequencia" runat="server" Width="110px" OnSelectedIndexChanged="ddlSequencia_SelectedIndexChanged"
                            AutoPostBack="True" />
                    </div>
                </div>
                <br />
                <div class="subtitulodiv">
                    TOTALIZADORES
                </div>
                <div class="row">
                    <div class="coltxt" style="text-align: center; width: 100%; line-height: 18px; font-weight: bold;">
                        Débitos
                        <br />
                        <asp:Label ID="lblTotalDebitos" runat="server" Style="color: Blue;" />
                        <br />
                        Créditos
                        <br />
                        <asp:Label ID="lblTotalCreditos" runat="server" Style="color: Red;" />
                        <br />
                        Diferença
                        <br />
                        <asp:Label ID="lblDiferenca" runat="server" Style="color: Green;" />
                    </div>
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaPlanoDeContas ID="ucConsultaPlanoDeContas" runat="server" />
    <uc:ConsultaPedidos ID="ucConsultaPedidos" runat="server" />
</asp:Content>
