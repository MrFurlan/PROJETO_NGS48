<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="AutorizacaoDeCarregamento.aspx.vb" Inherits="NGS.Web.UI.AutorizacaoDeCarregamento" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmAutorizacaoDeCarregamento" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlAutorizacaoDeCarregamento" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Autorização de Carregamento
            </div>
            <ajaxToolkit:TabContainer ID="TabContainer1" runat="server" ActiveTabIndex="0" Width="100%">
                <ajaxToolkit:TabPanel ID="TabPanel1" runat="server" HeaderText="TabPanel1">
                    <HeaderTemplate>
                        Cadastro
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes" runat="server">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconNovo" ID="lnkNovo" runat="server" Text="Gravar" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAtualizar" ID="lnkAtualizar" runat="server"
                                            Text="Atualizar" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconExcluir" ID="lnkExcluir" runat="server"
                                            Text="Excluir" OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconLimpar" ID="lnkLimpar" runat="server"
                                            Text="Limpar" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconRelatorio" ID="lnkRelatorio" runat="server"
                                            Text="Relatório" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAjuda" ID="lnkAjuda" runat="server"
                                            Text="Ajuda" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Placa:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtPlaca1" Enabled="False" runat="server" Width="95px" CssClass="texto" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnBuscaPlaca" runat="server" Text=">" CssClass="btn" />
                            </div>
                            <div class="collbl">
                                Motorista:
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="txtCodigoMotorista" runat="server" />
                                <asp:TextBox ID="txtMotorista" runat="server" CssClass="texto" Enabled="False" data-ToolTip="default"
                                    ToolTip="Dados do Motorista" Width="465px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                RNTRC:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtRNTRC1" Enabled="False" runat="server" CssClass="texto" MaxLength="8"
                                    Width="125px" />
                            </div>
                            <div class="collbl">
                                Transportador:
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="txtCodigoProprietario" runat="server" />
                                <asp:TextBox ID="txtNomeProprietario" Enabled="False" runat="server" Width="465px"
                                    MaxLength="60" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnBuscaProprietario" runat="server" Text=">" CssClass="btn" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Capacidade:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtCapacidade" Enabled="False" runat="server" Width="125px" MaxLength="60" />
                            </div>
                            <div class="collbl">
                                Tipo Veículo:
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="txtCodigoTipoVeiculo" runat="server" />
                                <asp:TextBox ID="txtDescricaoTipoVeiculo" Enabled="False" runat="server" Width="465px"
                                    MaxLength="60" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Pedido:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtPedido" Enabled="False" runat="server" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Entrega:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlEntregas" AutoPostBack="True" runat="server" Width="465px"
                                    MaxLength="60" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="subtitulodiv">
                                Roteiro(s)
                            </div>
                        </div>
                        <div class="row">
                            <div class="bordagrid" style="height: 115px;">
                                <asp:GridView ID="grdRoteiro" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                    ForeColor="#333333" GridLines="None" Width="100%" EmptyDataText="NENHUM REGISTRO ENCONTRADO"
                                    ShowHeaderWhenEmpty="True">
                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <EditRowStyle BackColor="#999999" />
                                    <EmptyDataRowStyle VerticalAlign="Bottom" HorizontalAlign="Center" />
                                    <Columns>
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
                                    </Columns>
                                </asp:GridView>
                            </div>
                        </div>
                        <div class="row">
                            <div class="painelleft" style="width: 40%;">
                                <div class="row">
                                    <div class="subtitulodiv">
                                        Produtos Autorizados
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="bordagrid" style="height: 115px;">
                                        <asp:GridView ID="grdProdutosAutorizados" runat="server" AutoGenerateColumns="False"
                                            CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%" EmptyDataText="NENHUM REGISTRO ENCONTRADO"
                                            ShowHeaderWhenEmpty="True">
                                            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                            <EditRowStyle BackColor="#999999" />
                                            <EmptyDataRowStyle VerticalAlign="Bottom" HorizontalAlign="Center" />
                                            <Columns>
                                                <asp:TemplateField>
                                                    <HeaderTemplate>
                                                        Cód. Produto
                                                    </HeaderTemplate>
                                                    <ItemTemplate>
                                                        <asp:Label ID="lblCodProd" runat="server" Text='<%# Eval("CodigoProduto") %>' />
                                                    </ItemTemplate>
                                                    <HeaderStyle HorizontalAlign="Left" />
                                                    <ItemStyle HorizontalAlign="Left" />
                                                </asp:TemplateField>
                                                <asp:TemplateField>
                                                    <HeaderTemplate>
                                                        Descrição
                                                    </HeaderTemplate>
                                                    <ItemTemplate>
                                                        <asp:Label ID="lblDescProd" runat="server" Text='<%# Eval("DescricaoProduto") %>' />
                                                    </ItemTemplate>
                                                    <HeaderStyle HorizontalAlign="Left" />
                                                    <ItemStyle HorizontalAlign="Left" />
                                                </asp:TemplateField>
                                                <asp:TemplateField>
                                                    <HeaderTemplate>
                                                        Qtde Autorizado
                                                    </HeaderTemplate>
                                                    <ItemTemplate>
                                                        <asp:Label ID="lblQtdeAut" runat="server" Text='<%# Eval("Quantidade") %>' />
                                                    </ItemTemplate>
                                                    <HeaderStyle HorizontalAlign="Left" />
                                                    <ItemStyle HorizontalAlign="Left" />
                                                </asp:TemplateField>
                                                <asp:TemplateField>
                                                    <ItemTemplate>
                                                        <asp:LinkButton ID="lnkAdicionaAutorizacao" runat="server" OnClick="lnkAdicionaAutorizacao_Click">
                                                            <asp:Image ID="imgAutCarregamento" runat="server" Width="16px" Height="16px" ImageUrl="~/Images/ico-mais.gif" />
                                                        </asp:LinkButton>
                                                    </ItemTemplate>
                                                    <HeaderStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                                    <ItemStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                                </asp:TemplateField>
                                            </Columns>
                                        </asp:GridView>
                                    </div>
                                </div>
                            </div>
                            <div class="painelright" style="width: 59%;">
                                <div class="row">
                                    <div class="subtitulodiv">
                                        Itens do Carregamento
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="bordagrid" style="height: 115px;">
                                        <asp:GridView ID="grdItensCarregamento" runat="server" AutoGenerateColumns="False"
                                            CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%" EmptyDataText="NENHUM REGISTRO ENCONTRADO"
                                            ShowHeaderWhenEmpty="True">
                                            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                            <EditRowStyle BackColor="#999999" />
                                            <EmptyDataRowStyle VerticalAlign="Bottom" HorizontalAlign="Center" />
                                            <Columns>
                                                <asp:CommandField SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button">
                                                    <ItemStyle Width="25px" />
                                                </asp:CommandField>
                                                <asp:TemplateField>
                                                    <HeaderTemplate>
                                                        Cód. Produto
                                                    </HeaderTemplate>
                                                    <ItemTemplate>
                                                        <asp:Label ID="lblCodProd" runat="server" Text='<%# Eval("CodigoProduto") %>' />
                                                    </ItemTemplate>
                                                    <HeaderStyle HorizontalAlign="Left" />
                                                    <ItemStyle HorizontalAlign="Left" />
                                                </asp:TemplateField>
                                                <asp:TemplateField>
                                                    <HeaderTemplate>
                                                        Descrição
                                                    </HeaderTemplate>
                                                    <ItemTemplate>
                                                        <asp:Label ID="lblDescProd" runat="server" Text='<%# Eval("Produto.Nome") %>' />
                                                    </ItemTemplate>
                                                    <HeaderStyle HorizontalAlign="Left" />
                                                    <ItemStyle HorizontalAlign="Left" />
                                                </asp:TemplateField>
                                                <asp:TemplateField HeaderText="Qtde Autorizado">
                                                    <ItemTemplate>
                                                        <asp:TextBox runat="server" ID="txtQtdeAutorizado" CssClass="txtDecimal4" AutoPostBack="true"
                                                            OnTextChanged="txtQtdeAutorizado_TextChanged" Text='<%# Eval("QuantidadeProgramado") %>' />
                                                    </ItemTemplate>
                                                    <HeaderStyle HorizontalAlign="Left" />
                                                    <ItemStyle HorizontalAlign="Left" />
                                                </asp:TemplateField>
                                                <asp:TemplateField HeaderText="% ADTO">
                                                    <ItemTemplate>
                                                        <asp:TextBox runat="server" ID="txtPec" CssClass="txtDecimal4" AutoPostBack="true"
                                                            OnTextChanged="txtPec_TextChanged" Text="0,0000" />
                                                    </ItemTemplate>
                                                    <HeaderStyle HorizontalAlign="Left" />
                                                    <ItemStyle HorizontalAlign="Left" />
                                                </asp:TemplateField>
                                                <asp:TemplateField>
                                                    <ItemTemplate>
                                                        <asp:ImageButton ID="btnExcluirItemCarregamento" runat="server" ImageUrl="~/Images/deletar.gif"
                                                            Style="border: 0;" data-ToolTip="default" ToolTip="Excluir" OnClientClick="return confirm('Confirma a exclusão deste item?');"
                                                            OnClick="btnExcluirItemCarregamento_Click" />
                                                    </ItemTemplate>
                                                    <HeaderStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                                    <ItemStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                                </asp:TemplateField>
                                            </Columns>
                                        </asp:GridView>
                                    </div>
                                </div>
                            </div>
                        </div>
                        <div class="row" runat="server">
                            <div class="painelleft" runat="server" id="tdNotaFiscal" visible="false" style="width: 100%;">
                                <div class="row">
                                    <div class="subtitulodiv">
                                        Notas Fiscais Vinculadas
                                        <asp:LinkButton ID="lnkNotaFiscal" runat="server" OnClick="lnkNotaFiscal_Click">
                                            <asp:Image ID="imgNotaFiscal" runat="server" Width="16px" margin-left="2px" Height="16px"
                                                ImageUrl="~/Images/ico-mais.gif" />
                                        </asp:LinkButton>
                                        <asp:Label ID="lblProdutoSelecionado" runat="server" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="bordagrid" style="height: 150px;">
                                        <asp:GridView ID="grdNotas" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                            ForeColor="#333333" GridLines="None" Width="100%" EmptyDataText="NENHUM REGISTRO ENCONTRADO"
                                            ShowHeaderWhenEmpty="True">
                                            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                            <EmptyDataRowStyle VerticalAlign="Bottom" HorizontalAlign="Center" />
                                            <EditRowStyle BackColor="#999999" />
                                            <Columns>
                                                <asp:TemplateField>
                                                    <HeaderTemplate>
                                                        Nota
                                                    </HeaderTemplate>
                                                    <ItemTemplate>
                                                        <asp:Label ID="lblNota" runat="server" Text='<%# Eval("Nota_Id") %>' />
                                                    </ItemTemplate>
                                                    <HeaderStyle HorizontalAlign="Left" />
                                                    <ItemStyle HorizontalAlign="Left" />
                                                </asp:TemplateField>
                                                <asp:TemplateField>
                                                    <HeaderTemplate>
                                                        Empresa
                                                    </HeaderTemplate>
                                                    <ItemTemplate>
                                                        <asp:Label ID="lblEmpresa" runat="server" Text='<%# Eval("Empresa_Id") %>' />
                                                    </ItemTemplate>
                                                    <HeaderStyle HorizontalAlign="Left" />
                                                    <ItemStyle HorizontalAlign="Left" />
                                                </asp:TemplateField>
                                                <asp:TemplateField>
                                                    <HeaderTemplate>
                                                        Cliente
                                                    </HeaderTemplate>
                                                    <ItemTemplate>
                                                        <asp:Label ID="lblCliente" runat="server" Text='<%# Eval("Cliente_Id") %>' />
                                                    </ItemTemplate>
                                                    <HeaderStyle HorizontalAlign="Left" />
                                                    <ItemStyle HorizontalAlign="Left" />
                                                </asp:TemplateField>
                                                <asp:TemplateField>
                                                    <HeaderTemplate>
                                                        E/S
                                                    </HeaderTemplate>
                                                    <ItemTemplate>
                                                        <asp:Label ID="lblEntradaSaida" runat="server" Text='<%# Eval("EntradaSaida_Id") %>' />
                                                    </ItemTemplate>
                                                    <HeaderStyle HorizontalAlign="Left" />
                                                    <ItemStyle HorizontalAlign="Left" />
                                                </asp:TemplateField>
                                                <asp:TemplateField>
                                                    <HeaderTemplate>
                                                        Produto
                                                    </HeaderTemplate>
                                                    <ItemTemplate>
                                                        <asp:Label ID="lblProduto" runat="server" Text='<%# Eval("Produto_Id") %>' />
                                                    </ItemTemplate>
                                                    <HeaderStyle HorizontalAlign="Left" />
                                                    <ItemStyle HorizontalAlign="Left" />
                                                </asp:TemplateField>
                                                <asp:TemplateField>
                                                    <HeaderTemplate>
                                                        CFOP
                                                    </HeaderTemplate>
                                                    <ItemTemplate>
                                                        <asp:Label ID="lblCFOP" runat="server" Text='<%# Eval("CFOP_Id") %>' />
                                                    </ItemTemplate>
                                                    <HeaderStyle HorizontalAlign="Left" />
                                                    <ItemStyle HorizontalAlign="Left" />
                                                </asp:TemplateField>
                                                <asp:TemplateField>
                                                    <HeaderTemplate>
                                                        Qtde Origem
                                                    </HeaderTemplate>
                                                    <ItemTemplate>
                                                        <asp:Label ID="lblQtde" runat="server" Text='<%# Eval("QuantidadeOrigem") %>' />
                                                    </ItemTemplate>
                                                    <HeaderStyle HorizontalAlign="Left" />
                                                    <ItemStyle HorizontalAlign="Left" />
                                                </asp:TemplateField>
                                                <asp:TemplateField>
                                                    <HeaderTemplate>
                                                        Qtde Destino
                                                    </HeaderTemplate>
                                                    <ItemTemplate>
                                                        <asp:Label ID="lblValor" runat="server" Text='<%# Eval("QuantidadeDestino") %>' />
                                                    </ItemTemplate>
                                                    <HeaderStyle HorizontalAlign="Left" />
                                                    <ItemStyle HorizontalAlign="Left" />
                                                </asp:TemplateField>
                                                <asp:TemplateField>
                                                    <ItemTemplate>
                                                        <asp:ImageButton ID="btnExcluirItemCarregamentoxNota" runat="server" ImageUrl="~/Images/deletar.gif"
                                                            Style="border: 0;" data-ToolTip="default" ToolTip="Excluir" OnClientClick="return confirm('Confirma a exclusão do vínculo com a nota?');"
                                                            OnClick="btnExcluirItemCarregamentoXNota_Click" />
                                                    </ItemTemplate>
                                                    <HeaderStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                                    <ItemStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                                </asp:TemplateField>
                                            </Columns>
                                        </asp:GridView>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel2" runat="server">
                    <HeaderTemplate>
                        Consultar
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="row">
                            <div class="collbl">
                                Carregamento:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtNroCarregamento" runat="server" CssClass="txtNumerico" Style="width: 100px;" />
                            </div>
                            <div class="collbl">
                                Placa:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtPlacaConsulta" Enabled="False" runat="server" Width="95px" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnConsultaPalca" runat="server" Text=">" CssClass="btn" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Transportador:
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="txtCodigoTranspConsulta" runat="server" />
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtTransportadorConsulta" Enabled="False" runat="server" Width="440px"
                                    MaxLength="60" />
                                <asp:Button ID="btnConsTransportador" runat="server" Text=">" CssClass="btn" />
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaPlacas ID="ucConsultaPlacas" runat="server" />
    <uc:ConsultaPedidos ID="ucConsultaPedidos" runat="server" />
    <uc:ConsultaNotaFiscal ID="ucConsultaNotaFiscal" runat="server" />
</asp:Content>
