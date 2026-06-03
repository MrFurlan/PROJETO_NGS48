<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="PesosDeChegada.aspx.vb" Inherits="NGS.Web.UI.PesosDeChegada" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngPesoDeChegada" runat="server" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlPesoDeChegada" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Pesos de Chegada
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconConsultar" runat="server">
                            <asp:LinkButton ID="lnkConsultar" runat="server" Text="Consultar">
                            </asp:LinkButton>
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" runat="server" Text="Limpar" />
                        </li>
                        <li class="iconRelatorio rel" runat="server" text="Relatório"><a>Relatório</a>
                            <ul>
                                <li class="iconPdf">
                                    <asp:LinkButton ID="lnkPdf" runat="server" Text="Pdf" />
                                </li>
                                <li class="iconExcel">
                                    <asp:LinkButton ID="lnkExcel" runat="server" Text="Excel" />
                                </li>
                            </ul>
                        </li>
                        <li class="iconAjuda" runat="server">
                            <asp:LinkButton ID="lnkAjuda" runat="server" Text="Ajuda" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="600px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                    <asp:TextBox ID="txtCliente" runat="server" Enabled="False" Width="560px" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnCliente" runat="server" Text=" > " OnClick="btnCliente_Click"
                        CssClass="btn" UseSubmitBehavior="False" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Pedido:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPedido" runat="server" Width="110px" Enabled="False" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnPedido" runat="server" Text=" > " OnClick="btnPedido_Click" UseSubmitBehavior="False"
                        CssClass="btn" />
                </div>
                <div class="collbl">
                    Número NF:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNotaFiscal" CssClass="txtNumerico9" runat="server" Width="150px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Transportador:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoTrans" runat="server" />
                    <asp:TextBox ID="txtTransportador" runat="server" Enabled="False" Width="560px" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnTransportador" runat="server" Text=" > " OnClick="btnTransportador_Click"
                        CssClass="btn" UseSubmitBehavior="False" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Grupo:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlGrupo" runat="server" Width="600px" OnSelectedIndexChanged="ddlGrupo_SelectedIndexChanged"
                        AutoPostBack="True" />
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
                    Período:
                </div>
                <div class="coltxt" style="width: 158px;">
                    <asp:TextBox ID="txtDataInicial" CssClass="calendario" runat="server" Width="80px" />
                </div>
                <div class="collbl">
                    à:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" CssClass="calendario" runat="server" Width="80px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    E/S:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdEntrada" runat="server" GroupName="entsai" Text=" Entrada " />
                </div>
                <div class="coltxt" style="width: 91px;">
                    <asp:RadioButton ID="rdSaida" runat="server" Checked="True" GroupName="entsai" Text=" Saída " />
                </div>
                <div class="collbl">
                    Tipo de Consulta:
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkViculados" runat="server" Text="Vinculados" />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="gridNF" runat="server" AutoGenerateColumns="False" CellPadding="1"
                    ForeColor="#333333" GridLines="None" width="100%">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:TemplateField>
                            <ItemTemplate>
                                <asp:LinkButton ID="lnkSelecionarNF" CssClass="lnk" 
                                    data-tooltip="default" ToolTip="Selecionar registro." runat="server" Text=" &gt; "
                                    OnClick="lnkSelecionarNF_Click">
                                    <i class="fa fa-arrow-right seta"></i>
                                </asp:LinkButton>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:BoundField DataField="EntradaSaida_Id" HeaderText="E/S">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Nota_Id" HeaderText="Nota">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Serie_Id" HeaderText="Ser">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Cliente_Id" HeaderText="Cliente">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="EndCliente_Id" HeaderText="End">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Nome" HeaderText="Nome">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Operacao" HeaderText="OP">
                            <HeaderStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Center" />
                        </asp:BoundField>
                        <asp:BoundField DataField="SubOperacao" HeaderText="SO">
                            <HeaderStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Center" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Data" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Data">
                            <HeaderStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Center" />
                        </asp:BoundField>
                        <asp:TemplateField HeaderText="Financeiro?">
                            <ItemTemplate>
                                <asp:Label ID="lblFinanceiro" runat="server" Text='<%# Eval("Financeiro")%>' />
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                        </asp:TemplateField>
                        <asp:BoundField DataField="PesoFiscal" DataFormatString="{0:N0}" HeaderText="Peso">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Movimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Movimento">
                            <HeaderStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Center" />
                        </asp:BoundField>
                        <asp:BoundField DataField="PesoBruto" HeaderText="Peso Bruto">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Desconto" HeaderText="Desconto">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField DataField="PesoLiquido" HeaderText="Peso Líquido">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Valor" DataFormatString="{0:N2}" HeaderText="Valor">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Placa" HeaderText="Placa">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField DataField="TarifaFrete" DataFormatString="{0:N2}" HeaderText="Tarifa Frete">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField DataField="ValorFrete" DataFormatString="{0:N2}" HeaderText="Valor Frete">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:TemplateField HeaderText=" ">
                            <ItemTemplate>
                                <asp:ImageButton ID="imgDelete" runat="server" ImageAlign="AbsMiddle" ImageUrl="~/Images/deletar.gif"
                                    Style="border: 0;" OnClick="imgDelete_Click" data-ToolTip="default" ToolTip="Excluir"
                                    OnClientClick="return confirm('Deseja realmente excluir este registro?');" />
                            </ItemTemplate>
                            <HeaderStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                            <ItemStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                        </asp:TemplateField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaPedidos ID="ucConsultaPedidos" runat="server" />
    <uc:PesoDeChegada ID="ucPesoDeChegada" runat="server" />
</asp:Content>
