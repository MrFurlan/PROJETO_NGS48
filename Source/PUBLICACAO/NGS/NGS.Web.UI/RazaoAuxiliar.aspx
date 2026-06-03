<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="RazaoAuxiliar.aspx.vb" Inherits="NGS.Web.UI.RazaoAuxiliar" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <script type="text/javascript">
        $(document).ready(function () {
            $("#txtDataInicial").setMask('date');
            $("#txtDataFinal").setMask('date');
            $("#txtDataEmissao").setMask('date');
        });

        function pageLoadRazaoAuxiliar() {
            $("div.accordion").accordion({
                active: false,
                collapsible: true,
                alwaysOpen: false,
                heightStyle: "content",
                autoHeight: false,
                clearStyle: true
            });

            if ($("#<%=chkConsolidarEmpresa.ClientID%>").attr("checked") == "checked")
                $("#divConsolidar").show();
            else
                $("#divConsolidar").hide();

            $("#<%=chkConsolidarEmpresa.ClientID%>").click(function () {
                if ($(this).attr("checked") == "checked")
                    $("#divConsolidar").show('drop');
                else
                    $("#divConsolidar").hide('drop');
            });
        }

        $(document).ready(function () {
            pageLoadRazaoAuxiliar();
            var prmRazaoAuxiliar = Sys.WebForms.PageRequestManager.getInstance();
            prmRazaoAuxiliar.add_endRequest(pageLoadRazaoAuxiliar);
        });
    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngRazaoAuxiliar" runat="server" AsyncPostBackTimeout="10000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlRazaoAuxiliar" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Razão Auxiliar
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconRelatorio rel" runat="server"><a>Relatório</a>
                            <ul>
                                <li class="iconPdf">
                                    <asp:LinkButton ID="lnkPdf" runat="server" Text="Pdf" />
                                </li>
                                <li class="iconExcel">
                                    <asp:LinkButton ID="lnkExcel" runat="server" Text="Excel" />
                                </li>
                            </ul>
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
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUnidadeDeNegocio" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlUnidadeDeNegocio_SelectedIndexChanged"
                        Width="590px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <div>
                        <asp:CheckBox ID="chkConsolidarEmpresa" runat="server" data-tooltip="default" title="Consolidar o cnpj da empresa. Modo Sintetico" Text="Empresa:" />
                    </div>

                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="590px" />
                </div>
            </div>
            <div id="divConsolidar" class="row">
                <div class="collbl">
                    Tipo Consolidação:
                </div>
                <div class="coltxt">

                    <asp:RadioButton ID="rbConsolidarEmpresaSintetico" GroupName="Consolidar" Checked="true" runat="server" data-ToolTip="default" ToolTip="Consolidar o cnpj da empresa. Modo Sintético." Text="Consolidação Sintética" />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rbConsolidarEmpresaAnalitico" GroupName="Consolidar" runat="server" data-ToolTip="default" ToolTip="Consolidar o cnpj da empresa. Modo Analítico." Text="Consolidação Analítica" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkCnpjCpf" runat="server" Text="Cliente" />
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNomeCliente" runat="server" Width="551px" Enabled="false" />
                    <asp:HiddenField ID="CodigoCliente" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="Button2" CssClass="btn" runat="server" Text=">" OnClick="btnCliente"
                        data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                </div>
            </div>
            <div class="row">
                <div class="coltxt" style="width: 100%; margin-bottom: 4px;">
                    <div id="divCustos" class="accordion" style="width: 100%;">
                        <h3>Seleção de C. Custo
                        </h3>
                        <asp:Panel ID="pnlCentroDeCusto" runat="server" Height="150px" ScrollBars="Vertical"
                            Width="94%">
                            <table style="width: 100%; border: 0px none;">
                                <tr>
                                    <td>
                                        <asp:CheckBoxList ID="SelecaoDeGrupos1" runat="server" />
                                    </td>
                                    <td>
                                        <asp:CheckBoxList ID="SelecaoDeGrupos2" runat="server" />
                                    </td>
                                </tr>
                            </table>
                        </asp:Panel>
                    </div>
                </div>
            </div>
            <%--            <div class="row">
                <div class="collbl">
                    Produto:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtProduto" runat="server" Width="551px" data-ToolTip="default"
                        ToolTip="Selecionar o produto." />
                    <asp:HiddenField ID="CodigoProduto" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnProduto" CssClass="btn" runat="server" OnClick="btnProduto_Click"
                        Text=">" data-ToolTip="default" ToolTip="Selecionar o produto." />
                </div>
            </div>--%>

            <div class="row" runat="server">
                <div class="coltxt" style="width: 100%; margin-bottom: 4px;">
                    <div class="coltxt" style="width: 100%;">
                        <uc:SelecaoProduto ID="ucSelecaoProduto" runat="server" />
                    </div>
                </div>
            </div>

            <div class="row">
                <div class="collbl">
                    Periodo:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicial" Width="90px" runat="server" ClientIDMode="Static"
                        CssClass="calendario" />
                </div>
                <div class="coltxt" style="width: 17px;">
                    À
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" runat="server" Width="90px" ClientIDMode="Static"
                        CssClass="calendario" data-ToolTip="default" />
                </div>
                <div class="collbl">
                    Emissao:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataEmissao" runat="server" Width="86px" ClientIDMode="Static"
                        CssClass="calendario" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Grupo de:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtGrupoInicial" Width="90px" runat="server">1</asp:TextBox>
                </div>
                <div class="coltxt">
                    <asp:ImageButton ID="ImgContaInicial" OnClick="ImgContaInicial_Click" runat="server"
                        ImageAlign="AbsMiddle" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png" />
                </div>
                <div class="coltxt" style="width: 19px;">
                    À
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtGrupoFinal" runat="server" Width="90px">4</asp:TextBox>
                </div>
                <div class="coltxt">
                    <asp:ImageButton ID="ImgContaFinal" OnClick="ImgContaFinal_Click" runat="server"
                        ImageAlign="AbsMiddle" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png" />
                </div>
                <div class="collbl">
                    Moeda:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlMoeda" runat="server" Width="96px">
                        <asp:ListItem>REAL</asp:ListItem>
                        <asp:ListItem>DOLAR</asp:ListItem>
                    </asp:DropDownList>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Listar:
                </div>
                <div class="coltxt" style="width: 268px;">
                    <asp:CheckBox ID="chkContasComMovimento" runat="server" Text="Apenas contas com movimento no período"
                        data-ToolTip="default" ToolTip="Selecionar para visualizar somente as contas com movimentação." />
                </div>
                <div class="collbl">
                    Iniciar na Folha:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtInicialFolha" runat="server" Width="86px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Lote:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtLote" runat="server" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Desconsiderar Lote:
                </div>
                <div class="coltxt" style="line-height: 16px;">
                    <asp:CheckBoxList ID="IsolarLotes" runat="server" Width="200px" data-ToolTip="default"
                        ToolTip="Marcar quando desejar desconsiderar lotes.">
                        <asp:ListItem Value="7000">7000 - Apuração De Custos</asp:ListItem>
                        <asp:ListItem Value="7001">7001 - Avaliação de Estoques</asp:ListItem>
                        <asp:ListItem Value="7500">7500 - Zeramento de Resultados</asp:ListItem>
                        <asp:ListItem Value="7600">7600 - Zeramento dos Estoques</asp:ListItem>
                    </asp:CheckBoxList>
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaProduto ID="ucConsultaProduto" runat="server" />
    <uc:ConsultaPlanoDeContas ID="ucConsultaPlanoDeContas" runat="server" />
</asp:Content>
