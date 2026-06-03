<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucNFEncargo.ascx.vb" Inherits="NGS.Web.UI.ucNFEncargo" %>

<div id="divNFEncargo" class="uc" title="Encargos do Produto" style="display: none;">

    <asp:UpdatePanel ID="updpnlNFEncargo" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <asp:HiddenField ID="hdfIndex" runat="server" />
            <asp:HiddenField ID="hdCentroDeCusto" runat="server" />
            <asp:HiddenField ID="observacaoValores" runat="server" />
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconConsultar" runat="server">
                            <asp:LinkButton ID="lnkRecarregar" Text="Recarregar" runat="server" />
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li class="iconFechar" runat="server">
                            <asp:LinkButton ID="lnkFechar" Text="Fechar" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Produto:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlProdutoSel" runat="server" Width="595px" AutoPostBack="true"
                        OnSelectedIndexChanged="ddlProdutoSel_SelectedIndexChanged" />
                </div>
            </div>
            <div class="row" runat="server" id="linhaBeneficio" visible="false">
                <div class="collbl">
                    Benefício Fiscal:
                </div>
                <div class="coltxt">
                    <asp:Label ID="txtBeneficioICMS" runat="server" BorderStyle="None" Font-Bold="False" ForeColor="Blue" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Conf. Operação:
                </div>
                <div class="texto_vermelho" style="width: 500px;">
                    Inicio Vigencia / Versão / Usuario / Data Inclusão
                </div>
            </div>
            <div class="row">
                <div style="white-space: nowrap; width: 113px; height: 26px; line-height: 26px; float: left; border-radius: 4px; font-family: Calibri; font-size: 12px; font-weight: bold; text-indent: 10px; margin-right: 4px; position: relative;">&nbsp;</div>
                <div>
                    &nbsp;&nbsp;<asp:DropDownList ID="ddlVersao" runat="server" Width="595px" Enabled="False" /> 
                </div>
            </div>


            <div class="bordagrid" style="height: auto;">

                <h3>Encargos da Nota:</h3>

                <asp:GridView ID="DgEncargos" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    OnRowCommand="DgEncargos_RowCommand"
                    OnRowCreated="DgEncargos_RowCreated"
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


            <div class="bordagrid" runat="server" id="idEncargosXml" visible="false" style="height: auto;">

                <h3>Encargos do Xml:</h3>

                <asp:GridView ID="DgEncargosXML" runat="server" AutoGenerateColumns="False" CellPadding="4"
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
</div>
