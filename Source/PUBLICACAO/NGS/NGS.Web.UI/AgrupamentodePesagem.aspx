<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="AgrupamentodePesagem.aspx.vb" Inherits="NGS.Web.UI.AgrupamentodePesagem" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <script type="text/javascript">
        function selectAll(chkAll) {
            var chk = $('#' + chkAll.id);
            var checked = chk.attr('checked') == "checked";

            $('#txtPrimeiraPesagem').val('0,00');
            $('#txtSegundaPesagem').val('0,00');
            $('#txtPesoBruto').val('0,00');
            $('#txtDesconto').val('0,00');
            $('#txtLiquido').val('0,00');

            $("input[type='checkbox']", "#MainContent_TabContainer1_TabPanel1_gridAgrupamentoDePesagem").not(".chkAlls").each(function () {
                $(this).attr("checked", checked);
                if (checked) {
                    SumTxt(this);
                }
            });
        };

        function SumTxt(ischk) {
            var chk = $('#' + ischk.id);
            var checked = chk.attr('checked') == "checked";
            var txt = $(ischk).parent().parent().find(".txtPrimeiraPesagem");

            if (txt.val() != undefined) {
                var PrimeiraPesagem = $(ischk).parent().parent().find(".txtPrimeiraPesagem").val().replace('.', '').replace(',', '.');
                var SegundaPesagem = $(ischk).parent().parent().find(".txtSegundaPesagem").val().replace('.', '').replace(',', '.');
                var BrutoBalanca = $(ischk).parent().parent().find(".txtBrutoBalanca").val().replace('.', '').replace(',', '.');
                var Desconto = $(ischk).parent().parent().find(".txtDesconto").val().replace('.', '').replace(',', '.');
                var Liquido = $(ischk).parent().parent().find(".txtLiquido").val().replace('.', '').replace(',', '.');

                var TotalPrimeiraPesagem = $('#txtPrimeiraPesagem').val().replace('.', '').replace(',', '.');
                var TotalSegunaPesagem = $('#txtSegundaPesagem').val().replace('.', '').replace(',', '.');
                var TotalBrutoBalanca = $('#txtPesoBruto').val().replace('.', '').replace(',', '.');
                var TotalDesconto = $('#txtDesconto').val().replace('.', '').replace(',', '.');
                var TotalLiquido = $('#txtLiquido').val().replace('.', '').replace(',', '.');

                if (checked) {
                    $('#txtPrimeiraPesagem').val((parseFloat(TotalPrimeiraPesagem) + parseFloat(PrimeiraPesagem)).toFixed(2).replace('.', ','));
                    $('#txtSegundaPesagem').val((parseFloat(TotalSegunaPesagem) + parseFloat(SegundaPesagem)).toFixed(2).replace('.', ','));
                    $('#txtPesoBruto').val((parseFloat(TotalBrutoBalanca) + parseFloat(BrutoBalanca)).toFixed(2).replace('.', ','));
                    $('#txtDesconto').val((parseFloat(TotalDesconto) + parseFloat(Desconto)).toFixed(2).replace('.', ','));
                    $('#txtLiquido').val((parseFloat(TotalLiquido) + parseFloat(Liquido)).toFixed(2).replace('.', ','));
                }
                else {
                    $('#txtPrimeiraPesagem').val((parseFloat(TotalPrimeiraPesagem) - parseFloat(PrimeiraPesagem)).toFixed(2).replace('.', ','));
                    $('#txtSegundaPesagem').val((parseFloat(TotalSegunaPesagem) - parseFloat(SegundaPesagem)).toFixed(2).replace('.', ','));
                    $('#txtPesoBruto').val((parseFloat(TotalBrutoBalanca) - parseFloat(BrutoBalanca)).toFixed(2).replace('.', ','));
                    $('#txtDesconto').val((parseFloat(TotalDesconto) - parseFloat(Desconto)).toFixed(2).replace('.', ','));
                    $('#txtLiquido').val((parseFloat(TotalLiquido) - parseFloat(Liquido)).toFixed(2).replace('.', ','));
                }
            }
        };

    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngAgrupamentoDePesagem" runat="server" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlAgrupamentoDePesagem" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Agrupamento de Pesagem
            </div>
            <ajaxToolkit:TabContainer Width="100%" ID="TabContainer1" runat="server" ActiveTabIndex="0"
                Style="margin-top: 4px;">
                <ajaxToolkit:TabPanel ID="TabPanel1" runat="server" HeaderText="TabPanel1">
                    <HeaderTemplate>
                        Agrupar
                    </HeaderTemplate>
                    <ContentTemplate>
                        <asp:Panel ID="pnlAgrupar" runat="server" Width="100%">
                            <div class="menu_acoes">
                                <div class="acoes">
                                    <ul>
                                        <li runat="server">
                                            <asp:LinkButton class="iconConsultar" ID="lnkConsultar" runat="server">
                                                <span>Consultar</span>
                                            </asp:LinkButton>
                                        </li>
                                        <li runat="server">
                                            <asp:LinkButton class="iconLimpar" ID="lnkLimpar" runat="server">
                                                <span>Limpar</span>
                                            </asp:LinkButton>
                                        </li>
                                        <li runat="server">
                                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" runat="server">
                                                <span>Ajuda</span>
                                            </asp:LinkButton>
                                        </li>
                                    </ul>
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Unidade:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlUnidadeDeNegocio" runat="server" Width="618px" AutoPostBack="True"
                                        OnSelectedIndexChanged="ddlUnidadeDeNegocio_SelectedIndexChanged" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Empresa:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="618px" AutoPostBack="True"
                                        OnSelectedIndexChanged="ddlEmpresa_SelectedIndexChanged" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Cliente:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtCliente" runat="server" Enabled="False" Width="579px" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnCliente" runat="server" Text=">" CssClass="btn" UseSubmitBehavior="False"
                                        OnClick="btnCliente_Click" data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Pedido:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtPedido" runat="server" Enabled="False" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnPedido" runat="server" Text=">" UseSubmitBehavior="False" OnClick="btnPedido_Click"
                                        CssClass="btn" data-ToolTip="default" ToolTip="Informar o número do pedido." />
                                    <asp:HiddenField ID="txtCodigoPedido" runat="server" />
                                </div>
                                <div class="collbl">
                                    Período:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtdata1agr" runat="server" Width="80px" CssClass="calendario" data-ToolTip="default"
                                        ToolTip="Data inicial a final para consulta." />
                                </div>
                                <div class="coltxt">
                                    á
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtdata2agr" runat="server" Width="80px" CssClass="calendario" data-ToolTip="default"
                                        ToolTip="Data inicial a final para consulta." />
                                </div>
                            </div>
                            <div class="bordagrid" style="height: 195px;">
                                <asp:GridView ID="gridAgrupamentoDePesagem" runat="server" CellPadding="3" ForeColor="#333333"
                                    GridLines="None" Width="100%" AutoGenerateColumns="False">
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <Columns>
                                        <asp:TemplateField>
                                            <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="5px" />
                                            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="5px" />
                                            <HeaderTemplate>
                                                <input id="chkAlls" onclick="selectAll(this);" type="checkbox" />
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <asp:CheckBox ID="chkPesagem" onclick="SumTxt(this);" runat="server" Checked='<%# eval("SelecionaPesagem") %>' />
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:BoundField DataField="Codigo" HeaderText="Pesagem">
                                            <HeaderStyle HorizontalAlign="Left" Width="80px" />
                                            <ItemStyle HorizontalAlign="Left" Width="80px" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Movimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Movimento">
                                            <HeaderStyle HorizontalAlign="Left" Width="80px" />
                                            <ItemStyle HorizontalAlign="Left" Width="80px" />
                                        </asp:BoundField>
                                        <asp:TemplateField>
                                            <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="70px" />
                                            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="70px" />
                                            <HeaderTemplate>
                                                1º Pesagem
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <asp:TextBox ID="txtPesagem1" Enabled="false" runat="server" Width="70px" CssClass="txtDecimal txtPrimeiraPesagem"
                                                    Style="text-align: right;" Text='<%# bind("PrimeiraPesagem", "{0:n2}") %>' />
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField>
                                            <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="70px" />
                                            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="70px" />
                                            <HeaderTemplate>
                                                2º Pesagem
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <asp:TextBox ID="txtPesagem2" Enabled="false" runat="server" Width="70px" CssClass="txtDecimal txtSegundaPesagem"
                                                    Style="text-align: right;" Text='<%# bind("SegundaPesagem", "{0:n2}") %>' />
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField>
                                            <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="70px" />
                                            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="70px" />
                                            <HeaderTemplate>
                                                Bruto
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <asp:TextBox ID="txtbbalanca" Enabled="false" runat="server" Width="70px" CssClass="txtDecimal txtBrutoBalanca"
                                                    Style="text-align: right;" Text='<%# bind("BrutoBalanca", "{0:n2}") %>' />
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField>
                                            <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="70px" />
                                            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="70px" />
                                            <HeaderTemplate>
                                                Desconto
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <asp:TextBox ID="txtdesc" Enabled="false" runat="server" Width="70px" CssClass="txtDecimal txtDesconto"
                                                    Style="text-align: right;" Text='<%# bind("Desconto", "{0:n2}") %>' />
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField>
                                            <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="70px" />
                                            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="70px" />
                                            <HeaderTemplate>
                                                Líquido
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <asp:TextBox ID="txtpesliquido" Enabled="false" runat="server" Width="70px" CssClass="txtDecimal txtLiquido"
                                                    Style="text-align: right;" Text='<%# bind("Liquido", "{0:n2}") %>' />
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField>
                                            <ItemTemplate>
                                                <asp:ImageButton ID="imgPesagem" runat="server" Height="20px" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png"
                                                    OnClick="imgPesagem_Click" data-ToolTip="default" ToolTip="Visualizar Pesagem"
                                                    Width="20px" />
                                            </ItemTemplate>
                                            <HeaderStyle Width="25px" />
                                            <ItemStyle Width="25px" />
                                        </asp:TemplateField>
                                    </Columns>
                                    <EditRowStyle BackColor="#999999" />
                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                </asp:GridView>
                            </div>
                            <div class="painelleft" style="width: 30%;">
                                <div class="row">
                                    <div class="collbl">
                                        1a. PESAGEM:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtPrimeiraPesagem" ClientIDMode="Static" CssClass="txtDecimal"
                                            Width="125px" runat="server" Enabled="False" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        2a. PESAGEM:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtSegundaPesagem" ClientIDMode="Static" runat="server" Enabled="False"
                                            CssClass="txtDecimal" Width="125px" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        PESO BRUTO:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtPesoBruto" ClientIDMode="Static" runat="server" Enabled="False"
                                            CssClass="txtDecimal" Width="125px" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        DESCONTO:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtDesconto" ClientIDMode="Static" runat="server" Enabled="False"
                                            CssClass="txtDecimal" Width="125px" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        LÍQUIDO:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtLiquido" ClientIDMode="Static" runat="server" Enabled="False"
                                            CssClass="txtDecimal" Width="125px" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="coltxt" style="float: right; position: relative; margin-right: 17px;">
                                        <asp:ImageButton ID="imgIncluir" runat="server" ImageUrl="~/images/confirmar.gif"
                                            CssClass="btn" data-ToolTip="default" ToolTip="Incluir Agrupamento" OnClick="imgIncluir_Click" />
                                    </div>
                                </div>
                            </div>
                            <div class="painelright" style="width: 70%;">
                                <div class="bordagrid" style="height: 200px;">
                                    <asp:GridView ID="gridDescontos" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                        ForeColor="#333333" GridLines="None" Width="100%">
                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                        <EditRowStyle BackColor="#999999" />
                                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                        <Columns>
                                            <asp:BoundField DataField="CodigoAnalise" HeaderText="Codigo">
                                                <HeaderStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="Descricao" HeaderText="Descri&#231;&#227;o">
                                                <HeaderStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                            <asp:TemplateField HeaderText="Percentual">
                                                <HeaderStyle HorizontalAlign="Left" />
                                                <ItemTemplate>
                                                    <asp:TextBox ID="txtPercentual" runat="server" Width="80px" Enabled="False" Text='<%# eval("Percentual") %>' />
                                                </ItemTemplate>
                                            </asp:TemplateField>
                                            <asp:TemplateField HeaderText="Indice">
                                                <HeaderStyle HorizontalAlign="Left" />
                                                <ItemTemplate>
                                                    <asp:TextBox ID="txtIndice" runat="server" ReadOnly="True" Width="56px" Enabled="False"
                                                        Text='<%# eval("Indice") %>' />
                                                </ItemTemplate>
                                            </asp:TemplateField>
                                            <asp:TemplateField HeaderText="Desconto">
                                                <HeaderStyle HorizontalAlign="Left" />
                                                <ItemTemplate>
                                                    <asp:TextBox ID="txtDesconto" runat="server" ReadOnly="True" Width="80px" Enabled="False"
                                                        Text='<%# eval("Desconto") %>' />
                                                </ItemTemplate>
                                            </asp:TemplateField>
                                        </Columns>
                                    </asp:GridView>
                                    <asp:HiddenField ID="hConfirmar" runat="server" />
                                    <asp:HiddenField ID="hAutorizacao" runat="server" />
                                </div>
                            </div>
                        </asp:Panel>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel2" runat="server" HeaderText="TabPanel2">
                    <HeaderTemplate>
                        Desagrupar
                    </HeaderTemplate>
                    <ContentTemplate>
                        <asp:Panel ID="pnlDesagrupar" runat="server" Width="100%">
                            <div class="menu_acoes">
                                <div class="acoes">
                                    <ul>
                                        <li class="iconConsultar" runat="server">
                                            <asp:LinkButton ID="lnkConsultarD" runat="server">
                                                                    <span>Consultar</span>
                                            </asp:LinkButton>
                                        </li>
                                        <li class="iconLimpar" runat="server">
                                            <asp:LinkButton ID="lnkLimparD" runat="server">
                                                                    <span>Limpar</span>
                                            </asp:LinkButton>
                                        </li>
                                    </ul>
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Unidade:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlUnidaDeNegocioDes" runat="server" Width="618px" AutoPostBack="True"
                                        OnSelectedIndexChanged="ddlUnidadeDeNegocioDes_SelectedIndexChanged" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Empresa:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlEmpresaDes" runat="server" Width="618px" AutoPostBack="True" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Cliente:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtClienteDes" runat="server" Enabled="False" Width="587px" />
                                    <asp:HiddenField ID="txtCodigoClienteDes" runat="server" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnClienteDes" CssClass="btn" runat="server" Text=">" UseSubmitBehavior="False" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Período:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtData1" CssClass="calendario" Width="75px" runat="server" data-ToolTip="default" ToolTip="Data inicial a final para consulta." />
                                    &nbsp;&nbsp;
                                    <asp:TextBox ID="txtData2" runat="server" Width="75px" CssClass="calendario" data-ToolTip="default" ToolTip="Data inicial a final para consulta." />
                                </div>
                                <div class="collbl">
                                    Pesagem:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtNumeroPesagem" CssClass="txtNumerico" runat="server" />
                                </div>
                            </div>
                            <div class="bordagrid">
                                <asp:GridView ID="grdDesagrupamento" runat="server" CellPadding="3" ForeColor="#333333"
                                    GridLines="None" Width="100%" AutoGenerateColumns="False">
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <Columns>
                                        <asp:BoundField DataField="Codigo" HeaderText="Pesagem">
                                            <HeaderStyle HorizontalAlign="Left" Width="80px" />
                                            <ItemStyle HorizontalAlign="Left" Width="80px" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Movimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Movimento">
                                            <HeaderStyle HorizontalAlign="Left" Width="80px" />
                                            <ItemStyle HorizontalAlign="Left" Width="80px" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Observacoes" HeaderText="Observacões">
                                            <HeaderStyle HorizontalAlign="Left" Width="200px" />
                                            <ItemStyle HorizontalAlign="Left" Width="200px" />
                                        </asp:BoundField>
                                        <asp:BoundField HeaderText="Bruto" DataField="BrutoBalanca" DataFormatString="{0:N0}">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                        <asp:BoundField HeaderText="Desconto" DataField="Desconto" DataFormatString="{0:N0}">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Liquido" HeaderText="L&#237;quido" DataFormatString="{0:N0}">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                        <asp:TemplateField>
                                            <ItemTemplate>
                                                <asp:ImageButton ID="imgDesfazerAgrupamento" runat="server" ImageUrl="~/images/deletar.gif"
                                                    CommandArgument='<%# Eval("Codigo") %>' Style="border: 0;" OnClick="imgDesfazerAgrupamento_Click"
                                                    data-ToolTip="default" ToolTip="Desfazer Agrupamento" OnClientClick="return confirm('Tem certeza que deseja Desagrupar o Laudo?');" />
                                            </ItemTemplate>
                                            <HeaderStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                            <ItemStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                        </asp:TemplateField>
                                    </Columns>
                                    <EditRowStyle BackColor="#999999" />
                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                </asp:GridView>
                            </div>
                        </asp:Panel>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaPedidos ID="ucConsultaPedidos" runat="server" />
</asp:Content>
