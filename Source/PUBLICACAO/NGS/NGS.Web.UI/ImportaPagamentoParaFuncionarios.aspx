<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master" CodeBehind="ImportaPagamentoParaFuncionarios.aspx.vb" Inherits="NGS.Web.UI.ImportaPagamentoParaFuncionarios" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        #meioconteudo {
            width: 1280px !important;
        }
    </style>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jquery.mask/1.14.16/jquery.mask.min.js"></script>
    <script type="text/javascript">
        function pageLoadImportaFolha() {
            $("#txtCompetencia:enabled").datepicker({
                dateFormat: 'mm/yy',
                monthNames: ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'],
                monthNamesShort: ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'],
                nextText: 'Próximo',
                prevText: 'Anterior',
                showOn: "button",
                buttonImage: "Images/calendar.png",
                buttonImageOnly: true,
                showButtonPanel: true
            });

            $("#txtCompetencia").mask("00/0000");
        }

        $(document).ready(function () {
            pageLoadImportaFolha();
            var prmImportaFolha = Sys.WebForms.PageRequestManager.getInstance();
            prmImportaFolha.add_endRequest(pageLoadImportaFolha);
        });
    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngImportaPagamentoParaFuncionarios" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlImportaPagamentoParaFuncionarios" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <ajaxToolkit:TabContainer ID="TabContainer1" runat="server" ActiveTabIndex="0" Width="100%">
                <ajaxToolkit:TabPanel ID="TabPanel1" runat="server" HeaderText="TabPanel1">
                    <HeaderTemplate>
                        Importação
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="titulodiv">
                            Importar Pagamento para Funcionários
                        </div>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconNovo" ID="lnkNovo" Text="Gravar" runat="server" ClientIDMode="Static" />
                                    </li>
                                    <li runat="server">
                                        <asp:UpdatePanel ID="updpnlImportar" runat="server">
                                            <ContentTemplate>
                                            <asp:LinkButton class="iconConsultar" ID="lnkImportar" Text="Importar" runat="server" />
                                            </ContentTemplate>
                                            <Triggers>
                                                <asp:PostBackTrigger ControlID="lnkImportar" />
                                            </Triggers>
                                        </asp:UpdatePanel>
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
                                <asp:DropDownList ID="ddlUnidadeNegocio" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlUnidadeNegocio_SelectedIndexChanged"
                                    Width="596px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Empresa:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlEmpresa" runat="server" Width="596px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Movimento:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtMovimnto" CssClass="calendario" runat="server" Width="96px" data-ToolTip="default" ToolTip="Informar o data para os Títulos a Pagar." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl"> 
                                Competência:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtCompetencia" runat="server" Width="96px" placeholder="MM/AAAA" ToolTip="Informar a competência da Folha." ClientIDMode="Static" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Tipos de Folha:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlTiposDeFolha" runat="server" Width="596px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Arquivo:
                            </div>
                            <div class="coltxt">
                                <input id="filUpload" style="width: 575px" type="file" runat="server" clientidmode="Static" />
                            </div>
                        </div>		
                        <div class="row" id="rowTotal" runat="server" visible="False">
                            <div class="coltxt">
                                <asp:Label ID="lblTotal" runat="server" Font-Bold="True" />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="gridFuncionarios" runat="server" AutoGenerateColumns="False"
                                CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:BoundField DataField="Empresa" HeaderText="Empresa">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Unidade" HeaderText="Unidade">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Movimento" HeaderText="Movimento" DataFormatString="{0:dd/MM/yyyy}">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Carteira" HeaderText="Carteira">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Encargo" HeaderText="Encargo">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="CPF" HeaderText="CPF">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Nome" HeaderText="Nome">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Banco" HeaderText="Banco">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Agencia" HeaderText="Agência">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="AgenciaDigito" HeaderText="Dig.">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Conta" HeaderText="Conta">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ContaDigito" HeaderText="Dig.">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Historico" HeaderText="Histórico">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Valor" HeaderText="Valor">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel2" runat="server" HeaderText="TabPanel2">
                    <HeaderTemplate>
                        Cancelamento
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="titulodiv">
                            Cancelar Pagamento para Funcionários
                        </div>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconNovo" ID="lnkCancelar" Text="Confirmar" runat="server" ClientIDMode="Static" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconConsultar" ID="lnkConsultarImportados" Text="Consultar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconLimpar" ID="lnkLimparImportados" Text="Limpar" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Unidade:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlUnidadeNegocio2" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlUnidadeNegocio2_SelectedIndexChanged"
                                    Width="596px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Empresa:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlEmpresa2" runat="server" Width="596px" />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="gridImportados" runat="server" AutoGenerateColumns="False"
                                CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:TemplateField>
                                        <HeaderTemplate>
                                            <asp:CheckBox ID="chkCancelarAll" data-ToolTip="default" ToolTip="Seleciona todos os títulos para cancelamento."
                                                Text="" runat="server" AutoPostBack="True" OnCheckedChanged="chkCancelarAll_CheckedChanged" />
                                        </HeaderTemplate>
                                        <ItemTemplate>
                                            <asp:CheckBox ID="chkCancelar" runat="server" data-ToolTip="default" ToolTip="Cancelar Registro" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="Codigo" HeaderText="Registro">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="CodigoEmpresa" HeaderText="Empresa">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Baixa" HeaderText="Movimento" DataFormatString="{0:dd/MM/yyyy}">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Carteira.Descricao" HeaderText="Carteira">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Tributo" HeaderText="Encargo">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="CodigoCliente" HeaderText="CPF">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Cliente.Nome" HeaderText="Nome">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="CodigoBancoCliente" HeaderText="Banco">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="CodigoAgenciaCliente" HeaderText="Agência">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="DigitoAgenciaCliente" HeaderText="Dig.">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ContaCliente" HeaderText="Conta">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="DigitoContaCliente" HeaderText="Dig.">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Historico" HeaderText="Histórico">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ValorDoDocumento" HeaderText="Valor">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
