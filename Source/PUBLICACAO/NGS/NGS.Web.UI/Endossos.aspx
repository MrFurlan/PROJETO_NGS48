<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="Endossos.aspx.vb" Inherits="NGS.Web.UI.Endossos" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        #meioconteudo {
            width: 1180px !important;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngEndossos" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlEndossos" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Endossos
            </div>
            <ajaxToolkit:TabContainer ID="TabEndossos" runat="server" ActiveTabIndex="0" Width="100%" Style="margin-top: 2px;">
                <ajaxToolkit:TabPanel ID="TabRegistro" runat="server">
                    <HeaderTemplate>
                        Lançamento
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconNovo" ID="lnkNovo" Text="Gravar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAtualizar" ID="lnkFinalizar" Text="Finalizar" runat="server"
                                            OnClientClick="if(!confirm('Deseja realmente finalizar o Endosso e liberar os Títuos?')) return false;" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconExcluir" ID="lnkExcluir" Text="Excluir" runat="server"
                                            OnClientClick="if(!confirm('Deseja realmente excluir o Endosso?')) return false;" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconConsultar" ID="lnkConsultar" Text="Consultar" runat="server" />
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
                                Código:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtSequencia" CssClass="txtNumerico" runat="server" Enabled="False" data-ToolTip="default"
                                    ToolTip="Sequência do Endosso." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Unidade:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlUnidadeNegocio" runat="server" Enabled="false" AutoPostBack="True" OnSelectedIndexChanged="ddlUnidadeNegocio_SelectedIndexChanged" Width="582px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Empresa:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlEmpresa" runat="server" Enabled="false" AutoPostBack="True" OnSelectedIndexChanged="ddlEmpresa_SelectedIndexChanged" Width="582px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Cliente:
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                                <asp:TextBox ID="txtCliente" runat="server" Width="543px" Enabled="False" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnCliente" runat="server" Enabled="false" Text=">" CssClass="btn" UseSubmitBehavior="False"
                                    data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Pedido:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtPedido" runat="server" Width="100px" Enabled="False" CssClass="txtNumerico" data-ToolTip="default" ToolTip="Número do pedido." />
                            </div>
                            <div class="coltxt">
                                <asp:Button CssClass="btn" ID="btnBuscaPedido" Enabled="false" OnClick="btnBuscaPedido_Click"
                                    runat="server" Text=">" CausesValidation="False" UseSubmitBehavior="False" data-ToolTip="default" ToolTip="Número do pedido." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Nota:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtNumNota" runat="server" Enabled="false" Width="339px" Style="text-align: right" data-ToolTip="default" ToolTip="Número da nota fiscal." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Data Inicial:
                            </div>
                            <div class="coltxt" style="width: 110px;">
                                <asp:TextBox ID="txtDataInicial" runat="server" Enabled="false" CssClass="calendario" Width="79px"
                                    data-ToolTip="default" ToolTip="Período final da pesquisa." />
                            </div>
                            <div class="collbl">
                                Data Final:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDataFinal" runat="server" Enabled="false" CssClass="calendario" Width="100px"
                                    data-ToolTip="default" ToolTip="Informar se a nota é de entrada ou saída." />
                            </div>
                        </div>
                        <div class="subtitulodiv">
                            <asp:Label ID="lblCabCliFor" runat="server" Text="Informações do Endosso" />
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Cliente:
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="txtCodigoClienteEndosso" runat="server" />
                                <asp:TextBox ID="txtClienteEndosso" runat="server" Width="543px" Enabled="False" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnClienteEndosso" runat="server" Enabled="false" Text=">" CssClass="btn" UseSubmitBehavior="False"
                                    data-ToolTip="default" ToolTip="Selecionar o Cliente para Endosso." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Número:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtNumEndosso" runat="server" Enabled="false" Style="text-align: right" data-ToolTip="default" ToolTip="Número do Endosso." />
                            </div>
                            <div class="collbl">
                                Vencimento:
                            </div>
                            <div class="coltxt" style="width: 110px;">
                                <asp:TextBox ID="txtVencimentoEndosso" runat="server" Enabled="false" CssClass="calendario" Width="79px"
                                    data-ToolTip="default" ToolTip="Vencimento do Endosso." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Valor:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="vlrEndosso" runat="server" Enabled="false" Style="text-align: right" data-ToolTip="default" CssClass="txtDecimal" ToolTip="Valor do Endosso." />
                            </div>
                            <div class="coltxt" id="rowValor" runat="server" visible="False">
                                <asp:Label ID="lblTotalRegistro" runat="server" Font-Bold="True" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Observação:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtObservacao" runat="server" Width="575px" TextMode="MultiLine" data-ToolTip="default" />
                            </div>
                        </div>
                        <div class="row" runat="server" visible="false" id="linhaGrid" >
                            <div class="bordagrid">
                                <asp:GridView ID="gridConsultaTitulos" runat="server" AutoGenerateColumns="False"
                                    CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <EditRowStyle BackColor="#999999" />
                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                    <Columns>
                                        <asp:TemplateField>
                                            <HeaderTemplate>
                                                <asp:CheckBox ID="chkAllTitulos" data-ToolTip="default" ToolTip="Seleciona todos os títulos de mesma moeda e indexador."
                                                    Text="CK" runat="server" AutoPostBack="True" OnCheckedChanged="chkAllTitulos_CheckedChanged" />
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <asp:CheckBox ID="chkGridTitulos" runat="server" AutoPostBack="True" OnCheckedChanged="chkGridTitulos_CheckedChanged" />
                                            </ItemTemplate>
                                            <HeaderStyle Width="30px" HorizontalAlign="Right" />
                                            <ItemStyle Width="30px" HorizontalAlign="Right" />
                                        </asp:TemplateField>
                                        <asp:BoundField DataField="Registro" HeaderText="Título">
                                            <HeaderStyle HorizontalAlign="Left" Width="60px" />
                                            <ItemStyle HorizontalAlign="Left" Width="60px" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Vencimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Vencimento" HtmlEncode="False">
                                            <HeaderStyle HorizontalAlign="Left" Width="80px" />
                                            <ItemStyle HorizontalAlign="Left" Width="80px" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Cliente" HeaderText="Cliente">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Historico" HeaderText="Histórico">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="ValorLiquido" DataFormatString="{0:N}" HeaderText="Reais" HtmlEncode="False">
                                            <HeaderStyle HorizontalAlign="Right" Width="100px" />
                                            <ItemStyle HorizontalAlign="Right" Width="100px" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Pedido" HeaderText="Pedido">
                                            <HeaderStyle HorizontalAlign="Center" Width="30px" />
                                            <ItemStyle HorizontalAlign="Center" Width="30px" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Grupado" HeaderText="Agrupado">
                                            <HeaderStyle HorizontalAlign="Center" Width="30px" />
                                            <ItemStyle HorizontalAlign="Center" Width="30px" />
                                        </asp:BoundField>
                                    </Columns>
                                </asp:GridView>
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabConsulta" runat="server">
                    <HeaderTemplate>
                        Consulta
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconConsultar" ID="lnkBusca" Text="Consultar" runat="server" />
                                    </li>
                                    <li class="iconRelatorio rel" runat="server"><a>Relatório</a>
                                        <ul>
                                            <li>
                                                <asp:LinkButton class="iconExcel" ID="lnkExcel" runat="server" Text="Excel" />
                                            </li>
                                        </ul>
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconLimpar" ID="lnkLimparBusca" Text="Limpar" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Código:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtSequenciaConsulta" CssClass="txtNumerico" runat="server" data-ToolTip="default"
                                    ToolTip="Sequência do Endosso." />
                            </div>
                            <div class="collbl">
                                Número:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtNumEndossoConsulta" runat="server" Style="text-align: right" data-ToolTip="default" ToolTip="Número do Endosso." />
                            </div>
                            <div class="collbl">
                                Vencimento:
                            </div>
                            <div class="coltxt" style="width: 110px;">
                                <asp:TextBox ID="txtVencimentoEndossoConsulta" runat="server" CssClass="calendario" Width="79px"
                                    data-ToolTip="default" ToolTip="Vencimento do Endosso." />
                            </div>
                            <div class="collbl">
                                Situação:
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="rbNormal" runat="server" Checked="True" GroupName="rdTipo"
                                    Text="Normal" data-ToolTip="default" ToolTip="Endosso Normal" AutoPostBack="True" />
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="rbFinalizado" runat="server" GroupName="rdTipo"
                                    Text="Finalizado" data-ToolTip="default" ToolTip="Endosso Finalizado" AutoPostBack="True" />
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="rbExcluido" runat="server" GroupName="rdTipo"
                                    Text="Excluído" data-ToolTip="default" ToolTip="Endosso Excluído" AutoPostBack="True" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Data Inicial:
                            </div>
                            <div class="coltxt" style="width: 110px;">
                                <asp:TextBox ID="txtBuscaDataInicial" runat="server" CssClass="calendario" Width="79px"
                                    data-ToolTip="default" ToolTip="Período inicial da pesquisa." />
                            </div>
                            <div class="collbl">
                                Data Final:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtBuscaDataFinal" runat="server" CssClass="calendario" Width="100px"
                                    data-ToolTip="default" ToolTip="Período final da pesquisa." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="bordagrid">
                                <asp:GridView ID="gridEndossos" runat="server" AutoGenerateColumns="False"
                                    CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="gridEndossos_SelectedIndexChanged">
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <EditRowStyle BackColor="#999999" />
                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                    <Columns>
                                        <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True">
                                            <HeaderStyle HorizontalAlign="Left" Width="25px" />
                                            <ItemStyle HorizontalAlign="Left" Width="25px" />
                                        </asp:CommandField>
                                        <asp:BoundField DataField="Codigo" HeaderText="Código">
                                            <HeaderStyle HorizontalAlign="Left" Width="60px" />
                                            <ItemStyle HorizontalAlign="Left" Width="60px" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Movimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Movimento" HtmlEncode="False">
                                            <HeaderStyle HorizontalAlign="Left" Width="80px" />
                                            <ItemStyle HorizontalAlign="Left" Width="80px" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="ClienteEndossoDescricao" HeaderText="Cliente do Endosso">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="NumeroEndosso" HeaderText="Número">
                                            <HeaderStyle HorizontalAlign="Left" Width="60px" />
                                            <ItemStyle HorizontalAlign="Left" Width="60px" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Vencimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Vencimento" HtmlEncode="False">
                                            <HeaderStyle HorizontalAlign="Left" Width="80px" />
                                            <ItemStyle HorizontalAlign="Left" Width="80px" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Valor" DataFormatString="{0:N}" HeaderText="Valor" HtmlEncode="False">
                                            <HeaderStyle HorizontalAlign="Right" Width="100px" />
                                            <ItemStyle HorizontalAlign="Right" Width="100px" />
                                        </asp:BoundField>
                                        <asp:TemplateField>
                                            <ItemTemplate>
                                                <asp:ImageButton ID="imgSituacao" runat="server" ImageUrl="~/Images/certo.jpg" Style="border: 0;" data-ToolTip="default" ToolTip="" />
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
                <ajaxToolkit:TabPanel ID="TabObservacao" runat="server">
                    <HeaderTemplate>
                        Observação
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="row">
                            <div class="collbl">
                                Controle Interno:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtObservacaoInterna" runat="server" Enabled="false" Width="1010px" Height="155px" TextMode="MultiLine" data-ToolTip="default" />
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaPedidos ID="ucConsultaPedidos" runat="server" />
</asp:Content>
