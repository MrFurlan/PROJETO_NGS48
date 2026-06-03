<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="EnvioDeXML.aspx.vb" Inherits="NGS.Web.UI.EnvioDeXML" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        #meioconteudo {
            width: 1360px !important;
        }

        .txtNumerico {
            text-align: right;
        }

        .tt-titulos .tooltip-inner {
            max-width: 520px;
            white-space: normal;
            font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", monospace;
            text-align: left;
            background-color: #fff; /* fundo branco */
            color: #000; /* texto preto para contraste */
            border: 1px solid #ccc; /* borda leve para destacar */
            box-shadow: 0 2px 6px rgba(0,0,0,.15); /* sombra suave */
        }

        .badge {
            display: inline-block;
            padding: .35em .6em;
            border-radius: .375rem;
            font-weight: 600;
            line-height: 1;
        }

        .status-0 {
            background: #dc3545;
            color: #fff;
        }

        .status-1 {
            background: #ffc107;
            color: #000;
        }

        .status-2 {
            background: #198754;
            color: #fff;
        }

        .legenda-status {
            display: flex;
            justify-content: flex-end; /* empurra para direita */
            gap: 1.5rem; /* espaço entre cada item */
            margin-top: 5px; /* distância do grid */
            width: 100%; /* ocupa a mesma largura do grid */
        }

            .legenda-status span {
                display: inline-flex;
                align-items: center;
                font-size: 0.9rem;
            }

            .legenda-status .badge {
                margin-right: 6px;
            }

        .tt-icon {
            cursor: pointer;
        }
    </style>

    <script type="text/javascript">
        (function () {
            function ensureBootstrap(cb) {
                if (window.bootstrap && bootstrap.Tooltip) { cb(); return; }
                var s = document.createElement('script');
                s.src = 'https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js';
                s.onload = cb;
                document.head.appendChild(s);
            }

            window.initTooltips = function () {
                ensureBootstrap(function () {
                    var els = [].slice.call(document.querySelectorAll('[data-bs-toggle="tooltip"]'));
                    els.forEach(function (el) {
                        try { var inst = bootstrap.Tooltip.getInstance(el); if (inst) inst.dispose(); } catch (e) { }
                        new bootstrap.Tooltip(el, { html: true, sanitize: false, customClass: 'tt-titulos' });
                    });
                });
            };

            document.addEventListener("DOMContentLoaded", initTooltips);

            if (window.Sys && Sys.WebForms && Sys.WebForms.PageRequestManager) {
                Sys.WebForms.PageRequestManager.getInstance().add_endRequest(function () { initTooltips(); });
            }
        })();
    </script>
</asp:Content>


<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmEnvioDeXML" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlEnvioDeXML" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Envio de XML
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconMail" ID="lnkEnviarEmail" runat="server" Text="Enviar E-mail" />
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
                    Pedido:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPedido" Enabled="false" runat="server" Width="100px" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="cmdBuscaPedido" OnClick="cmdBuscaPedido_Click" runat="server" Text=">"
                        CssClass="btn" CausesValidation="False" UseSubmitBehavior="False" data-ToolTip="default"
                        ToolTip="Número do pedido." />
                </div>
            </div>
            <div class="row">
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
                    Documento:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdNFe" runat="server" Text="Nota Fiscal" AutoPostBack="True" Checked="True" GroupName="Documento" data-ToolTip="default" ToolTip="Nota Fiscal Eletrônica." />
                    <asp:RadioButton ID="rdCTe" runat="server" Text="Conhecimento de Transporte" AutoPostBack="True" GroupName="Documento" data-ToolTip="default" ToolTip="Conhecimento de Trasporte Eletrônico." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Tipo:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdEntradas" runat="server" Text="Entradas" AutoPostBack="True" Checked="True" GroupName="Notas" data-ToolTip="default" ToolTip="Apenas nota de entrada." />
                    <asp:RadioButton ID="rdSaidas" runat="server" Text="Saidas" AutoPostBack="True" GroupName="Notas" data-ToolTip="default" ToolTip="Apenas nota de saída." />
                    <asp:RadioButton ID="rdTodas" runat="server" Text="Todas" AutoPostBack="True" GroupName="Notas" data-ToolTip="default" ToolTip="Todas as Notas." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Modo:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdConferencia" runat="server" Text="Conferência" AutoPostBack="True" Checked="True" GroupName="Tipo" data-ToolTip="default" ToolTip="Envio de XML para conferência." />
                    <asp:RadioButton ID="rdBanco" runat="server" Text="Envio ao Banco" Enabled="false" AutoPostBack="True" GroupName="Tipo" data-ToolTip="default" ToolTip="Envio de Xml para Instituições Bancárias." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Anexar PDF:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdPDFNao" runat="server" Text="Não" Checked="True" GroupName="EnviarPDF" data-ToolTip="default" ToolTip="Não Anexar PDF." />
                    <asp:RadioButton ID="rdPDFSim" runat="server" Text="Sim" GroupName="EnviarPDF" data-ToolTip="default" ToolTip="Anexar PDF." />
                </div>
                <div class="collbl">
                    Compactado:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdZIPNao" runat="server" Text="Não" Checked="True" GroupName="Zipar" data-ToolTip="default" ToolTip="Não enviar compactado." />
                    <asp:RadioButton ID="rdZIPSim" runat="server" Text="Sim" GroupName="Zipar" data-ToolTip="default" ToolTip="Enviar compactado." />
                </div>
                <div class="coltxt legenda-status" id="tipoTitulo" runat="server" visible="false">
                    <span><span class="badge status-2">&#9679;</span>Liberado</span>
                    <span><span class="badge status-1">&#9679;</span>Parcialmente Liberado</span>
                    <span><span class="badge status-0">&#9679;</span>Vencido</span>
                </div>
            </div>
            <div class="bordagrid" style="min-height: 400px; max-height: 400px; max-width: 100%; width: 100%; overflow: auto;">
                <asp:GridView ID="gridDocumentos" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:TemplateField>
                            <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="5px" />
                            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="5px" />
                            <HeaderTemplate>
                                <asp:CheckBox ID="chkEnviarAll" data-ToolTip="default" ToolTip="Seleciona todas as Notas."
                                    Text="" runat="server" AutoPostBack="True" OnCheckedChanged="chkEnviarAll_CheckedChanged" />
                            </HeaderTemplate>
                            <ItemTemplate>
                                <asp:CheckBox ID="chkEnviar" runat="server" data-ToolTip="default" ToolTip="Selecionar Nota" AutoPostBack="True" OnCheckedChanged="chkEnviar_CheckedChanged" />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:BoundField DataField="Movimento" HeaderText="Movimento" DataFormatString="{0:dd/MM/yyyy}">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Empresa_id" HeaderText="Empresa">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="EndEmpresa_id" HeaderText="End">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Cliente" HeaderText="Cliente" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="EndCliente" HeaderText="End">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="NomeCliente" HeaderText="Nome">
                            <HeaderStyle HorizontalAlign="Left" Width="350px" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="ES" HeaderText="E/S">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Serie" HeaderText="S&#233;rie">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Nota" HeaderText="Nota">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Quantidade" HeaderText="Quantidade" DataFormatString="{0:n0}">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" Width="100px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Valor" HeaderText="Valor" DataFormatString="{0:n2}">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" Width="100px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="ChaveNfe" HeaderText="">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:TemplateField HeaderText="Títulos">
                            <HeaderStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Center" />
                            <ItemTemplate>
                                <asp:Image ID="imgTitulos" runat="server"
                                    ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png"
                                    ImageAlign="AbsMiddle"
                                    AlternateText="Ver título"
                                    ToolTip='<%# TooltipHistorico(Eval("Historico")) %>'
                                    CssClass="tt-icon" />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Situação">
                            <HeaderStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Center" />
                            <ItemTemplate>
                                <span class='<%# "badge status-" & (If(Eval("Situacao") Is DBNull.Value, "2", Eval("Situacao"))) %>'>&#9679;</span>
                                </span>
                            </ItemTemplate>
                        </asp:TemplateField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaPedidos ID="ucConsultaPedidos" runat="server" />
    <uc:EmailNFe ID="ucEmailNFe" runat="server" />
</asp:Content>
