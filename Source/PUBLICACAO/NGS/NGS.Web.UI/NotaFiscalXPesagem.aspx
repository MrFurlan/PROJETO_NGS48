<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="NotaFiscalXPesagem.aspx.vb" Inherits="NGS.Web.UI.NotaFiscalXPesagem" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngNotaFiscalXPesagem" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlNotaFiscalXPesagem" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                NotaFiscal X Pesagem
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconConsultar" ID="lnkConsultar" runat="server" Text="Gravar" />
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
                    <asp:TextBox ID="txtCliente" runat="server" Enabled="False" Width="569px" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnCliente" runat="server" Text=">" OnClick="btnCliente_Click" UseSubmitBehavior="False"
                        CssClass="btn" data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
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
                    <asp:Button ID="btnPedido" runat="server" Text=">" OnClick="btnPedido_Click" UseSubmitBehavior="False"
                        CssClass="btn" data-ToolTip="default" ToolTip="Número do pedido. " />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdEntrada" runat="server" GroupName="entsai" Text="Entrada"
                        data-ToolTip="default" ToolTip="Inforamar se é entrada ou saída." />
                    <asp:RadioButton ID="rdSaida" runat="server" Checked="True" GroupName="entsai" Text="Saída"
                        data-ToolTip="default" ToolTip="Inforamar se é entrada ou saída." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicial" CssClass="calendario" runat="server" Width="80px"
                        data-ToolTip="default" ToolTip="Período da consulta." />
                </div>
                <div class="collbl">
                    à:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" CssClass="calendario" runat="server" Width="80px"
                        data-ToolTip="default" ToolTip="Período da consulta." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Nota Fiscal:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNotaFiscal" runat="server" Enabled="False" data-ToolTip="default"
                        ToolTip="Número da nota fiscal." />
                </div>
                <div class="collbl">
                    Pesagem:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPesagem" runat="server" Enabled="False" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnPesagem" runat="server" Text=">" Enabled="False" OnClick="btnPesagem_Click"
                        UseSubmitBehavior="False" CssClass="btn" data-ToolTip="default" ToolTip="Peso do produto." />
                </div>
                <div class="coltxt">
                    <asp:ImageButton ID="imgConfirmar" runat="server" ImageUrl="~/images/confirmar.gif"
                        ImageAlign="AbsMiddle" data-ToolTip="default" ToolTip="Confirmar vinculo do Laudo"
                        OnClick="imgConfirmar_Click" />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="gridNF" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
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
                        <asp:BoundField DataField="PesoFiscal" DataFormatString="{0:N0}" HeaderText="Peso">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Valor" DataFormatString="{0:N2}" HeaderText="Valor">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Romaneio_Id" HeaderText="Romaneio">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaPedidos ID="ucConsultaPedidos" runat="server" />
    <uc:ConsultaRomaneios ID="ucConsultaRomaneios" runat="server" />
</asp:Content>
