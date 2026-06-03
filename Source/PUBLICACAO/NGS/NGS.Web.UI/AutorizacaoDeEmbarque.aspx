<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="AutorizacaoDeEmbarque.aspx.vb" Inherits="NGS.Web.UI.AutorizacaoDeEmbarque" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <%-- <script type="text/javascript">
        function validarFrete(txt) {
            var ctrl = $('#' + txt.id);
            if ($.trim(ctrl.val()) != "" && $.trim(ctrl.val()) != "0,00" && ctrl.val() != undefined) {
                ctrl.parent().parent().find(".txtValorTonelada").val("0,00");
            }
        }

        function validarTonelada(txt) {
            var ctrl = $('#' + txt.id);
            if ($.trim(ctrl.val()) != "" && $.trim(ctrl.val()) != "0,00" && ctrl.val() != undefined) {
                ctrl.parent().parent().find(".txtValorFret").val("0,00");
            }
        }

        $(document).ready(function () {
            pageLoadAutorizacaoDeEmbarque();
            var prmAutorizacaoDeEmbarque = Sys.WebForms.PageRequestManager.getInstance();
            prmAutorizacaoDeEmbarque.add_endRequest(pageLoadAutorizacaoDeEmbarque);
        });
    </script>--%>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmAutorizacaoDeEmbarque" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="uppnlAutorizacaoDeEmbarque" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Autorização de Embarque
            </div>
            <ajaxToolkit:TabContainer ID="TabContainer1" runat="server" ActiveTabIndex="0" Width="100%">
                <ajaxToolkit:TabPanel ID="TabPanel1" runat="server">
                    <HeaderTemplate>
                        Pedidos
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconConsultar" ID="lnkConsultarNovo" Text="Consultar" runat="server" />
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
                                Empresa:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlEmpresa" runat="server" Width="635px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Cliente:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtNomeCliente" Enabled="False" runat="server" Width="595px" />
                                <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnCliente" runat="server" CssClass="btn" Text="&gt;" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Produto:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtNomeProduto" Enabled="False" runat="server" Width="595px" />
                                <asp:HiddenField ID="CodigoProduto" runat="server" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnProduto" runat="server" Text="&gt;" CssClass="btn" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Safra:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlSafra" runat="server" Width="438px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Pedido:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtPedido" runat="server" Width="100px" />
                            </div>
                            <div class="collbl">
                                Tipo de Operação:
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="rdTodos" runat="server" Checked="True" GroupName="ES" Text="Todos" />
                                <asp:RadioButton ID="rdEntrada" runat="server" GroupName="ES" Text="Entrada" />
                                <asp:RadioButton ID="rdSaida" runat="server" GroupName="ES" Text="Saida" />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="grdPedidos" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                EmptyDataText="NENHUM REGISTRO ENCONTRADO" ForeColor="#333333" GridLines="None"
                                OnSelectedIndexChanged="grdPedidos_SelectedIndexChanged" ShowHeaderWhenEmpty="True"
                                Width="100%">
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <Columns>
                                    <asp:CommandField ButtonType="Button" SelectText="&gt;" ShowSelectButton="True" />
                                    <asp:BoundField DataField="CodigoPedido" HeaderText="Pedido" />
                                    <asp:BoundField DataField="DescOperacao" HeaderText="Operacao" />
                                    <asp:BoundField DataField="CodigoCliente" HeaderText="Cliente" />
                                    <asp:BoundField DataField="EndCliente" HeaderText="End" />
                                    <asp:BoundField DataField="NomeCliente" HeaderText="Nome Cliente" />
                                    <asp:BoundField DataField="CodigoProduto" HeaderText="Produto" />
                                    <asp:BoundField DataField="DataEntrega" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Data Entrega" />
                                    <asp:TemplateField FooterText="Situação" HeaderText="Situação" ShowHeader="False">
                                        <ItemTemplate>
                                            <asp:Button ID="BtnAlteraSituacaoPedido" runat="server" CausesValidation="False"
                                                Font-Bold="True" OnClick="BtnAlteraSituacaoPedido_Click" Style="cursor: pointer;"
                                                Text='<%# bind("EmbarqueAtivo") %>' Width="90px" />
                                        </ItemTemplate>
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
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel2" runat="server">
                    <HeaderTemplate>
                        Autorização
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li class="iconRelatorio" runat="server">
                                        <asp:LinkButton ID="lnkRelatorio" Text="Relatório" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="subtitulodiv">
                            DADOS DO PEDIDO
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Safra:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtSafraAut" Enabled="False" runat="server" Width="180px" />
                            </div>
                            <div class="collbl">
                                Empresa:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtEmpresaAut" Enabled="False" runat="server" Width="471px" />
                                <asp:HiddenField ID="txtCodigoEmpresaAut" runat="server" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Pedido:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtPedidoAut" Enabled="False" runat="server" Width="180px" />
                            </div>
                            <div class="collbl">
                                Cliente:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtClienteAut" Enabled="False" runat="server" Width="471px" />
                                <asp:HiddenField ID="txtCodigoClienteAut" runat="server" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Data Entrega:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDataEntregaAut" Enabled="False" runat="server" Width="180px" />
                            </div>
                            <div class="collbl">
                                Operação:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtOperacao" Enabled="False" runat="server" Width="471px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Produto:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtProduto" Width="180px" runat="server" Enabled="false" Font-Bold="true" />
                            </div>
                            <div class="collbl">
                                Cfop:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtCfop" Enabled="False" runat="server" Width="471px" />
                                <asp:Label ID="d" runat="server" Font-Bold="True" />
                            </div>
                        </div>
                        <ajaxToolkit:TabContainer ID="TabContConfigPedido" runat="server" ActiveTabIndex="0"
                            Width="100%">
                            <ajaxToolkit:TabPanel ID="TabLocaisDeEntrega" runat="server">
                                <HeaderTemplate>
                                    Locais de Entrega
                                </HeaderTemplate>
                                <ContentTemplate>
                                    <div class="subtitulodiv">
                                        LOCAIS DE ENTREGA
                                    </div>
                                    <div class="menu_acoes">
                                        <div class="acoes">
                                            <ul>
                                                <li class="iconNovo" runat="server" style="width: 14%;">
                                                    <asp:LinkButton ID="lnkNovoLocalEntregaClientePedido" Text="Adicionar Local" runat="server" />
                                                </li>
                                            </ul>
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="collbl">
                                            Local de Entrega:
                                        </div>
                                        <div class="coltxt">
                                            <asp:TextBox ID="txtNomeLocaldeEntrega" Enabled="False" runat="server" Width="475px" />
                                            <asp:HiddenField ID="txtCodigoLocalEntrega" runat="server" />
                                        </div>
                                        <div class="coltxt">
                                            <asp:Button ID="btnConsultaLocalEntrega" Text=" > " runat="server" CssClass="btn" />
                                        </div>
                                        <div class="coltxt">
                                            <asp:CheckBox ID="chkContaEOrdem" runat="server" Text="Emitir Nota de Conta e Ordem" />
                                        </div>
                                    </div>
                                    <div class="bordagrid" style="height: 215px;">
                                        <asp:GridView ID="gridLocalEntrega" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                            ForeColor="#333333" GridLines="None" Width="100%">
                                            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                            <EditRowStyle BackColor="#999999" />
                                            <Columns>
                                                <asp:CommandField SelectText="&gt;" ShowSelectButton="True" ButtonType="Button">
                                                    <ItemStyle Width="25px" />
                                                </asp:CommandField>
                                                <asp:CheckBoxField DataField="EmitirNota" HeaderText="Emitir Nota">
                                                    <HeaderStyle HorizontalAlign="Center" />
                                                    <ItemStyle HorizontalAlign="Center" />
                                                </asp:CheckBoxField>
                                                <asp:TemplateField>
                                                    <HeaderTemplate>
                                                        CNPJ
                                                    </HeaderTemplate>
                                                    <ItemTemplate>
                                                        <asp:Label ID="lblCodigoClienteEntrega" runat="server" Text='<%# Eval("CodigoClienteEntrega") %>' />
                                                    </ItemTemplate>
                                                    <HeaderStyle HorizontalAlign="Left" />
                                                    <ItemStyle HorizontalAlign="Left" />
                                                </asp:TemplateField>
                                                <asp:TemplateField>
                                                    <HeaderTemplate>
                                                        Entrega
                                                    </HeaderTemplate>
                                                    <ItemTemplate>
                                                        <asp:Label ID="lblEntregaRes" runat="server" Text='<%# Eval("DescClienteEntrega") %>' />
                                                    </ItemTemplate>
                                                    <HeaderStyle HorizontalAlign="Left" />
                                                    <ItemStyle HorizontalAlign="Left" />
                                                </asp:TemplateField>
                                                <asp:TemplateField>
                                                    <HeaderTemplate>
                                                        Qtde Autorizado
                                                    </HeaderTemplate>
                                                    <ItemTemplate>
                                                        <asp:Label ID="lblNormal" runat="server" Text='<%# Bind("Produtos.QtdeAutorizadoLE","{0:N2}") %>' />
                                                    </ItemTemplate>
                                                    <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                                    <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                                </asp:TemplateField>
                                                <asp:TemplateField HeaderText="Aut. +">
                                                    <HeaderTemplate>
                                                        Aut. +
                                                    </HeaderTemplate>
                                                    <ItemTemplate>
                                                        <asp:LinkButton ID="lnkAutEmbarque" runat="server" OnClick="lnkAutEmbarque_Click">
                                                            <asp:Image ID="imgAutEmbarque" data-ToolTip="default" ToolTip="Adiciona ou extorna quantidades autorizadas"
                                                                runat="server" Width="16px" Height="16px" ImageUrl="~/Images/ico-mais.gif" />
                                                        </asp:LinkButton>
                                                    </ItemTemplate>
                                                    <ItemStyle HorizontalAlign="Center" />
                                                </asp:TemplateField>
                                                <asp:TemplateField HeaderText="Observação">
                                                    <ItemTemplate>
                                                        <asp:LinkButton ID="lnkAdicionarObservacao" data-ToolTip="default" ToolTip="Adiona observações ao local de entrega."
                                                            runat="server" OnClick="lnkAdicionarObservacao_Click">
                                                            <asp:Image runat="server" ImageUrl="~/Images/ico-mais.gif" Height="16px" Width="16px"
                                                                ID="imgObservacao"></asp:Image>
                                                        </asp:LinkButton>
                                                    </ItemTemplate>
                                                    <ItemStyle HorizontalAlign="Center" />
                                                </asp:TemplateField>
                                                <asp:TemplateField>
                                                    <HeaderTemplate>
                                                        Excluir
                                                    </HeaderTemplate>
                                                    <ItemTemplate>
                                                        <asp:ImageButton ID="btnExcluirAutorizacao" runat="server" ImageUrl="~/Images/deletar.gif"
                                                            OnClick="btnExcluirAutorizacao_Click" OnClientClick="return confirm('Atencão! /n Ao Excluir este local de Entrega todas as autorizacoes vinculadas a este local serão excluidos /n Deseja realmente excluir este Local de Entrega?');"
                                                            Style="border: 0;" data-ToolTip="default" ToolTip="Excluir" />
                                                    </ItemTemplate>
                                                    <ItemStyle HorizontalAlign="Center" />
                                                </asp:TemplateField>
                                            </Columns>
                                        </asp:GridView>
                                    </div>
                                </ContentTemplate>
                            </ajaxToolkit:TabPanel>
                            <ajaxToolkit:TabPanel ID="TabRoteiros" runat="server">
                                <HeaderTemplate>
                                    Roteiros
                                </HeaderTemplate>
                                <ContentTemplate>
                                    <div id="pnlRoteiro" runat="server">
                                        <div class="menu_acoes">
                                            <div class="acoes">
                                                <ul>
                                                    <li class="iconNovo" runat="server" style="width: 16%;">
                                                        <asp:LinkButton ID="lnkAdicionarRoteiro" Text="Adicionar Roteiro" runat="server" />
                                                    </li>
                                                </ul>
                                            </div>
                                        </div>
                                        <div class="subtitulodiv">
                                            ROTEIROS:
                                            <div style="float: right;">
                                                <asp:Label ID="lblRecebeLocalDeEntrega" runat="server" />
                                            </div>
                                        </div>
                                        <div class="bordagrid" style="height: 115px;">
                                            <asp:GridView ID="grdRoteiro" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                                ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="GridRoteiro_SelectedIndexChanged">
                                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                <EditRowStyle BackColor="#999999" />
                                                <Columns>
                                                    <asp:CommandField SelectText="&gt;" ShowSelectButton="True" ButtonType="Button">
                                                        <ItemStyle Width="25px" />
                                                    </asp:CommandField>
                                                    <asp:TemplateField>
                                                        <HeaderTemplate>
                                                            Origem
                                                        </HeaderTemplate>
                                                        <ItemTemplate>
                                                            <asp:Label ID="lblNomeOrigem" runat="server" Text='<%# Eval("OrigemDescricao") %>' />
                                                        </ItemTemplate>
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:TemplateField>
                                                    <asp:TemplateField>
                                                        <HeaderTemplate>
                                                            Destino
                                                        </HeaderTemplate>
                                                        <ItemTemplate>
                                                            <asp:Label ID="lblNomeDestino" runat="server" Text='<%# Eval("DestinoDescricao") %>' />
                                                        </ItemTemplate>
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:TemplateField>
                                                    <asp:TemplateField>
                                                        <HeaderTemplate>
                                                            Via de Transporte
                                                        </HeaderTemplate>
                                                        <ItemTemplate>
                                                            <asp:Label ID="lblViaTransporte" runat="server" Text='<%# Eval("ViaDeTransporte.Descricao") %>' />
                                                        </ItemTemplate>
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:TemplateField>
                                                    <asp:TemplateField HeaderText="Excluir">
                                                        <ItemTemplate>
                                                            <asp:ImageButton ID="btnExcluirRoteiro" runat="server" ImageUrl="~/Images/deletar.gif"
                                                                CommandArgument='<%# Eval("CodigoRoteiro") %>' Style="border: 0;" OnClick="btnExcluirRoteiro_Click"
                                                                data-ToolTip="default" ToolTip="Excluir" OnClientClick="return confirm('Atencão!/r/nAo Excluir este roteiro todos os Transportadores e Preços vinculados a ele serão excluidos/r/nDeseja realmente excluir este roteiro?');" />
                                                        </ItemTemplate>
                                                        <HeaderStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                                        <ItemStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                                    </asp:TemplateField>
                                                </Columns>
                                            </asp:GridView>
                                        </div>
                                    </div>
                                </ContentTemplate>
                            </ajaxToolkit:TabPanel>
                            <ajaxToolkit:TabPanel ID="TabTransportadores" runat="server">
                                <HeaderTemplate>
                                    Transportadores/Preços
                                </HeaderTemplate>
                                <ContentTemplate>
                                    <div id="pnlTransportador" runat="server" visible="False">
                                        <div class="subtitulodiv">
                                            Roteiro Selecionado:
                                            <div style="float: right;">
                                                <asp:Label ID="lblRoteiroDe" runat="server" />
                                            </div>
                                        </div>
                                        <div class="subtitulodiv" style="text-align: right;">
                                            <asp:Label ID="lblRoteiroPara" runat="server" />
                                        </div>
                                        <div class="menu_acoes">
                                            <div class="acoes">
                                                <ul>
                                                    <li class="iconNovo" runat="server" style="width: 22.1%;">
                                                        <asp:LinkButton ID="lnkNovoTransportador" Text="Adicionar Transportador" runat="server" />
                                                    </li>
                                                </ul>
                                            </div>
                                        </div>
                                        <div class="row">
                                            <div class="collbl">
                                                Transportador:
                                            </div>
                                            <div class="coltxt">
                                                <asp:TextBox ID="txtNomeTransportador" runat="server" Enabled="False" Width="358px" />
                                                <asp:HiddenField ID="txtCodigoTransportador" runat="server" />
                                            </div>
                                            <div class="coltxt">
                                                <asp:Button ID="btnConsultaTransportador" runat="server" Text="&gt;" CssClass="btn" />
                                            </div>
                                            <div class="collbl" style="width: 80px;">
                                                Quota:
                                            </div>
                                            <div class="coltxt">
                                                <asp:TextBox ID="txtQuota" CssClass="txtDecimal4" runat="server" Width="80px" />
                                                <asp:Label ID="lblUnidade" runat="server" Font-Bold="True" />
                                            </div>
                                            <div class="coltxt">
                                                <asp:DropDownList ID="ddlPesoQuantidade" runat="server" AutoPostBack="True" Style="width: 140px;">
                                                    <asp:ListItem Value="P">Peso Do Produto</asp:ListItem>
                                                    <asp:ListItem Value="Q">Numero de Fretes</asp:ListItem>
                                                </asp:DropDownList>
                                            </div>
                                        </div>
                                        <div class="bordagrid" style="height: 119px;">
                                            <asp:GridView ID="grdTransportador" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                                ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="GridTransportador_SelectedIndexChanged">
                                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                <EditRowStyle BackColor="#999999" />
                                                <Columns>
                                                    <asp:CommandField SelectText="&gt;" ShowSelectButton="True" ButtonType="Button">
                                                        <ItemStyle Width="25px" />
                                                    </asp:CommandField>
                                                    <asp:BoundField DataField="CnpjTransp" HeaderText="Código">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>
                                                    <asp:TemplateField HeaderText="Nome">
                                                        <HeaderTemplate>
                                                            Nome
                                                        </HeaderTemplate>
                                                        <ItemTemplate>
                                                            <asp:Label ID="lblNomeTransportador" runat="server" Text='<%# Eval("Transportador.Nome") %>' />
                                                        </ItemTemplate>
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:TemplateField>
                                                    <asp:BoundField DataField="Quota" DataFormatString="{0:N4}" HeaderText="Quota">
                                                        <HeaderStyle HorizontalAlign="Right" />
                                                        <ItemStyle HorizontalAlign="Right" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="DescPesoQuantidade" HeaderText="Unid.">
                                                        <HeaderStyle HorizontalAlign="Right" />
                                                        <ItemStyle HorizontalAlign="Right" />
                                                    </asp:BoundField>
                                                    <asp:TemplateField FooterText="Ativo" HeaderText="Ativo" ShowHeader="False">
                                                        <ItemTemplate>
                                                            <asp:Button ID="btnAlteraAtivo" runat="server" CausesValidation="False" Font-Bold="True"
                                                                OnClick="btnAlteraAtivo_Click" Style="cursor: pointer;" Text='<%# bind("Ativo") %>'
                                                                Width="90px" />
                                                        </ItemTemplate>
                                                        <HeaderStyle Width="100px" HorizontalAlign="Center" />
                                                        <ItemStyle Width="100px" HorizontalAlign="Center" />
                                                    </asp:TemplateField>
                                                    <asp:TemplateField HeaderText="Excluir">
                                                        <ItemTemplate>
                                                            <asp:ImageButton ID="btnExcluirTransportador" runat="server" ImageUrl="~/Images/deletar.gif"
                                                                CommandArgument='<%# Eval("CnpjTransp") %>' Style="border: 0;" OnClick="btnExcluirTransportador_Click"
                                                                data-ToolTip="default" ToolTip="Excluir" OnClientClick="return confirm('Atencão \n\nAo Excluir este Transportador, todos os Preços a ele vinculados serão excluidos. \nDeseja realmente excluir este transportador?');" />
                                                        </ItemTemplate>
                                                        <HeaderStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                                        <ItemStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                                    </asp:TemplateField>
                                                </Columns>
                                            </asp:GridView>
                                        </div>
                                    </div>
                                    <div id="pnlPreco" runat="server" visible="false">
                                        <div class="subtitulodiv">
                                            Preços
                                            <div style="float: right;">
                                                <asp:Label ID="lblRecebeTransp" runat="server" />
                                            </div>
                                        </div>
                                        <div class="menu_acoes">
                                            <div class="acoes">
                                                <ul>
                                                    <li class="iconNovo" runat="server" style="width: 16%;">
                                                        <asp:LinkButton ID="lnkGravarValores" Text="Adicionar Preço" runat="server" />
                                                    </li>
                                                </ul>
                                            </div>
                                        </div>
                                        <div class="row">
                                            <div class="collbl">
                                                Quota:
                                            </div>
                                            <div class="coltxt">
                                                <asp:TextBox ID="txtQuotaPrecoTransportador" CssClass="txtDecimal" runat="server" />
                                            </div>
                                            <div class="collbl">
                                                Valor:
                                            </div>
                                            <div class="coltxt">
                                                <asp:TextBox ID="txtValorPrecoTransportador" CssClass="txtDecimal" runat="server" />
                                            </div>
                                        </div>
                                        <div class="bordagrid" style="height: 115px;">
                                            <asp:GridView ID="grdPrecos" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                                ForeColor="#333333" GridLines="None" Width="100%">
                                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                <EditRowStyle BackColor="#999999" />
                                                <Columns>
                                                    <asp:BoundField DataField="nrCotacao" DataFormatString="{0:N0}" HeaderText="Cotação">
                                                        <HeaderStyle HorizontalAlign="Right" />
                                                        <ItemStyle HorizontalAlign="Right" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="Movimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Data">
                                                        <HeaderStyle HorizontalAlign="Right" />
                                                        <ItemStyle HorizontalAlign="Right" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="Quota" DataFormatString="{0:N4}" HeaderText="Quota">
                                                        <HeaderStyle HorizontalAlign="Right" />
                                                        <ItemStyle HorizontalAlign="Right" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="ValorFrete" DataFormatString="{0:N2}" HeaderText="Valor Frete">
                                                        <HeaderStyle HorizontalAlign="Right" />
                                                        <ItemStyle HorizontalAlign="Right" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="ValorTon" DataFormatString="{0:N2}" HeaderText="Valor Ton.">
                                                        <HeaderStyle HorizontalAlign="Right" />
                                                        <ItemStyle HorizontalAlign="Right" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="UsuarioInclusao" HeaderText="Usuario" ReadOnly="True"
                                                        SortExpression="UsuarioInclusao" />
                                                    <asp:TemplateField HeaderText="Excluir">
                                                        <ItemTemplate>
                                                            <asp:ImageButton ID="btnExcluirPreco" runat="server" ImageUrl="~/Images/deletar.gif"
                                                                Style="border: 0;" OnClick="btnExcluirPreco_Click" data-ToolTip="default" ToolTip="Excluir"
                                                                OnClientClick="return confirm('Deseja realmente excluir este Preço?');" />
                                                        </ItemTemplate>
                                                        <HeaderStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                                        <ItemStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                                    </asp:TemplateField>
                                                </Columns>
                                            </asp:GridView>
                                        </div>
                                    </div>
                                </ContentTemplate>
                            </ajaxToolkit:TabPanel>
                        </ajaxToolkit:TabContainer>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:AutorizacaoDeEmbarque ID="ucAutorizacaoDeEmbarque" runat="server" />
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaProduto ID="ucConsultaProduto" runat="server" />
    <uc:OrigemDestino ID="ucOrigemDestino" runat="server" />
    <uc:ConsultaObservacoes ID="ucConsultaObservacoes" runat="server" />
    <uc:ConsultaObservacoesEmbarque ID="ucConsultaObservacoesEmbarque" runat="server" />
</asp:Content>
