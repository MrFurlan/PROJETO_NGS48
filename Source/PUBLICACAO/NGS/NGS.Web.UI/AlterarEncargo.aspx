<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="AlterarEncargo.aspx.vb" Inherits="NGS.Web.UI.AlterarEncargo" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        #meioconteudo {
            width: 1360px !important;
        }

        .colw {
            width: 165px;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngAlterarEncargo" runat="server" AsyncPostBackTimeout="1000">
    </asp:ScriptManager>
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlAlterarEncargo" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <asp:HiddenField ID="idGridItem" runat="server" />
            <asp:HiddenField ID="hdCentroDeCusto" runat="server" />
            <div class="titulodiv">
                Alterar Encargos
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconAtualizar" ID="lnkAtualizar" Text="Atualizar" runat="server"
                                OnClientClick="if(!confirm('Deseja realmente alterar este registro?')) return false;" />
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
                <div class="collbl colw">
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUnidadeDeNegocio" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlUnidadeDeNegocio_SelectedIndexChanged"
                        Width="650px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl colw">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="650px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl colw">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                    <asp:TextBox ID="txtCliente" runat="server" Enabled="False" Width="611px" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnCliente" runat="server" OnClick="btnCliente_Click" CssClass="btn"
                        Text=">" UseSubmitBehavior="False" data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl colw">
                    Data Inicial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicial" runat="server" CssClass="calendario" Width="86px"
                        data-ToolTip="default" ToolTip="Período inicial da pesquisa." />
                </div>
                <div class="collbl">
                    Data Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" runat="server" CausesValidation="True" CssClass="calendario"
                        Width="86px" data-ToolTip="default" ToolTip="Período final da pesquisa." />
                </div>
            </div>
            <div class="row">
                <div class="collbl colw">
                    Nota Fiscal:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtES" runat="server" MaxLength="1" Width="50px" Enabled="true" placeholder="E/S"
                        data-ToolTip="default" ToolTip="Informações da nota fiscal (Tipo, número e série)." />
                    <asp:TextBox ID="txtNota" runat="server" CssClass="txtNumerico" Enabled="true" placeholder="Nota"
                        data-ToolTip="default" ToolTip="Informações da nota fiscal (Tipo, número e série)." />
                    <asp:TextBox ID="txtSerie" runat="server" MaxLength="3" Width="50px" Enabled="true" placeholder="Série"
                        data-ToolTip="default" ToolTip="Informações da nota fiscal (Tipo, número e série)." />
                </div>
            </div>
            <div class="bordagrid" style="height: 150px;">
                <asp:GridView ID="gridItens" runat="server" Width="100%" ForeColor="#333333" OnSelectedIndexChanged="gridItens_SelectedIndexChanged" GridLines="None" CellPadding="4" AutoGenerateColumns="False">
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775"></AlternatingRowStyle>
                    <Columns>
                        <asp:CommandField InsertText="" SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button"></asp:CommandField>
                        <asp:BoundField DataField="CFOP" HeaderText="CFOP">
                            <HeaderStyle HorizontalAlign="Right" Width="30px" />
                            <ItemStyle HorizontalAlign="Right" Width="30px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="CodigoProduto" HeaderText="Produto">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Produto.Nome" HeaderText="Nome Produto">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="QuantidadeFisica" DataFormatString="{0:N4}" HeaderText="Qtde. Física">
                            <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="QuantidadeFiscal" DataFormatString="{0:N4}" HeaderText="Qtde. Fiscal">
                            <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Unitario" DataFormatString="{0:N10}" HeaderText="Unitário">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField DataField="ValorTotal" DataFormatString="{0:N2}" HeaderText="Valor">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                    </Columns>
                    <EditRowStyle BackColor="#999999"></EditRowStyle>
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White"></FooterStyle>
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White"></HeaderStyle>
                    <PagerStyle HorizontalAlign="Center" BackColor="#284775" ForeColor="White"></PagerStyle>
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333"></RowStyle>
                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333"></SelectedRowStyle>
                </asp:GridView>
            </div>
             <div class="subtitulodiv" id="idSubtitulo" runat="server" visible="false">
                   <asp:Label ID="lblProduto" runat="server" Font-Bold="True" Font-Size="12px" />
                </div>
            <div class="row" runat="server" id="idBeneficio" visible="false">
                <div class="collbl">
                    Benefício Fiscal:
                </div>
                <div class="coltxt">
                    <asp:Label ID="txtBeneficioICMS" runat="server" BorderStyle="None" Font-Bold="False" ForeColor="Blue" />
                </div>
            </div>
            <div class="row" id="idVigencia" runat="server" visible="false">
                <div class="collbl">
                    Conf. Operação:
                </div>
                <div class="texto_vermelho" style="width: 500px;">
                    Inicio Vigencia / Versão / Usuario / Data Inclusão
                </div>
            </div>
            <div class="row" id="idVersao" runat="server" visible="false">
                <div style="white-space: nowrap; width: 113px; height: 26px; line-height: 26px; float: left; border-radius: 4px; font-family: Calibri; font-size: 12px; font-weight: bold; text-indent: 10px; margin-right: 4px; position: relative;">
                    &nbsp;
                </div>
               
                <div>
                    &nbsp;&nbsp;<asp:DropDownList ID="ddlVersao" runat="server" Width="595px" Enabled="False" />
                </div>
            </div>
            <div class="bordagrid" id="tabEncargos" runat="server" visible="false" style="height: auto;">
                <asp:GridView ID="gridEncargos" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />

                    <Columns>
                        <asp:BoundField DataField="Codigo" HeaderText="Encargo">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="SituacaoTributaria" HeaderText="ClaICMS">
                            <HeaderStyle HorizontalAlign="right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="right"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="SituacaoTributariaPISCOFINS" HeaderText="PisCofins">
                            <HeaderStyle HorizontalAlign="right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="right"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="SituacaoTributariaIPI" HeaderText="IPI">
                            <HeaderStyle HorizontalAlign="right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="right"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="CodigoOperacao" HeaderText="Op">
                            <HeaderStyle HorizontalAlign="right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="right"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="CodigoSubOperacao" HeaderText="So">
                            <HeaderStyle HorizontalAlign="right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="right"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="CodigoGrupoProduto" HeaderText="Grupo">
                            <HeaderStyle HorizontalAlign="right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="right"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="CodigoProduto" HeaderText="Produto">
                            <HeaderStyle HorizontalAlign="right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="right"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="EstadoOrigem" HeaderText="Origem">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="EstadoDestino" HeaderText="Destino">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                        </asp:BoundField>
                        <asp:TemplateField HeaderText="Base">
                            <ItemTemplate>
                                <asp:TextBox ID="txtBaseEncargoItem" runat="server" CssClass="txtDecimal" BorderStyle="None"
                                    Enabled="False" Text='<%# Eval("Base", "{0:N2}") %>' Width="100px" />
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="right" />
                            <ItemStyle HorizontalAlign="right" />
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Percentual">
                            <ItemTemplate>
                                <asp:TextBox ID="txtPercentualEncargoItem" runat="server" CssClass="txtDecimal9"
                                    BorderStyle="None" Enabled="False" Text='<%# Eval("Percentual", "{0:N9}") %>'
                                    Width="100px" />
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="right" />
                            <ItemStyle HorizontalAlign="right" />
                        </asp:TemplateField>
                        <asp:BoundField DataField="PercentualExibicao" DataFormatString="{0:N9}" HeaderText="% Exib.">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:TemplateField HeaderText="Valor">
                            <ItemTemplate>
                                <asp:TextBox ID="txtValorEncargoItem" runat="server" CssClass="txtDecimal" Enabled="False"
                                    Text='<%# Eval("Valor", "{0:N2}") %>' BorderStyle="None" Width="100px" />
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="right"></ItemStyle>
                        </asp:TemplateField>
                        <asp:BoundField DataField="Sinal" HeaderText="Sinal">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
            <div class="row" runat="server" id="idBeneficioNovo" visible="false">
                <div class="collbl">
                    Benefício Fiscal:
                </div>
                <div class="coltxt">
                    <asp:Label ID="txtBeneficioICMSNovo" runat="server" BorderStyle="None" Font-Bold="False" ForeColor="Blue" />
                </div>
            </div>
            <div class="row" id="idVigenciaNova" runat="server" visible="false">
                <div class="collbl">
                    Conf. Operação:
                </div>
                <div class="texto_vermelho" style="width: 500px;">
                    Inicio Vigencia / Versão / Usuario / Data Inclusão
                </div>
            </div>
            <div class="row" id="idVersaoNova" runat="server" visible="false">
                <div style="white-space: nowrap; width: 113px; height: 26px; line-height: 26px; float: left; border-radius: 4px; font-family: Calibri; font-size: 12px; font-weight: bold; text-indent: 10px; margin-right: 4px; position: relative;">
                    <asp:Button ID="btnRecarregar" OnClick="btnRecarregar_Click" runat="server" CssClass="botao" UseSubmitBehavior="False" Text="Recarregar"></asp:Button>
                </div>
                <div>
                    &nbsp;&nbsp;<asp:DropDownList ID="ddlVersaoNova" runat="server" Width="595px" Enabled="False" />
                </div>
            </div>
            <div class="bordagrid" id="tabEncargosNovos" runat="server" visible="false" style="height: auto;">
                <asp:GridView ID="gridEncargosNovos" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    OnRowCommand="gridEncargosNovos_RowCommand"
                    OnRowCreated="gridEncargosNovos_RowCreated"
                    ForeColor="#333333" GridLines="None" Width="100%">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />

                    <Columns>
                        <asp:BoundField DataField="Codigo" HeaderText="Encargo">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="SituacaoTributaria" HeaderText="ClaICMS">
                            <HeaderStyle HorizontalAlign="right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="right"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="SituacaoTributariaPISCOFINS" HeaderText="PisCofins">
                            <HeaderStyle HorizontalAlign="right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="right"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="SituacaoTributariaIPI" HeaderText="IPI">
                            <HeaderStyle HorizontalAlign="right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="right"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="CodigoOperacao" HeaderText="Op">
                            <HeaderStyle HorizontalAlign="right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="right"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="CodigoSubOperacao" HeaderText="So">
                            <HeaderStyle HorizontalAlign="right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="right"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="CodigoGrupoProduto" HeaderText="Grupo">
                            <HeaderStyle HorizontalAlign="right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="right"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="CodigoProduto" HeaderText="Produto">
                            <HeaderStyle HorizontalAlign="right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="right"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="EstadoOrigem" HeaderText="Origem">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="EstadoDestino" HeaderText="Destino">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                        </asp:BoundField>
                        <asp:TemplateField HeaderText="Base">
                            <ItemTemplate>
                                <asp:TextBox ID="txtBaseEncargoItem" runat="server" CssClass="txtDecimal" BorderStyle="None"
                                    Enabled="False" Text='<%# Eval("Base", "{0:N2}") %>' Width="100px" />
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="right" />
                            <ItemStyle HorizontalAlign="right" />
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Percentual">
                            <ItemTemplate>
                                <asp:TextBox ID="txtPercentualEncargoItem" runat="server" CssClass="txtDecimal9"
                                    BorderStyle="None" Enabled="False" Text='<%# Eval("Percentual", "{0:N9}") %>'
                                    Width="100px" />
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="right" />
                            <ItemStyle HorizontalAlign="right" />
                        </asp:TemplateField>
                        <asp:BoundField DataField="PercentualExibicao" DataFormatString="{0:N9}" HeaderText="% Exib.">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:TemplateField HeaderText="Valor">
                            <ItemTemplate>
                                <asp:TextBox ID="txtValorEncargoItem" runat="server" CssClass="txtDecimal" Enabled="False"
                                    Text='<%# Eval("Valor", "{0:N2}") %>' BorderStyle="None" Width="100px" />
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="right"></ItemStyle>
                        </asp:TemplateField>
                        <asp:BoundField DataField="Sinal" HeaderText="Sinal">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                        </asp:BoundField>
                        <asp:TemplateField>
                            <ItemTemplate>
                                <asp:Button ID="btnEncargoItem" runat="server" CommandName="OK" UseSubmitBehavior="False"
                                    Enabled="False" Text="OK" />
                            </ItemTemplate>
                        </asp:TemplateField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaPedidosxNotas ID="ucConsultaPedidosXNotas" runat="server" />
</asp:Content>
