<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucConsultaDadosBancarios.ascx.vb"
    Inherits="NGS.Web.UI.ucConsultaDadosBancarios" %>
<div id="divConsultaDadosBancarios" class="uc" title="Consulta de Dados Bancários" style="display: none;">
    <asp:UpdatePanel ID="updpnlConsultaDadosBancarios" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="row" style="line-height: 14px;">
                <ajaxToolkit:TabContainer ID="TabContainer1" runat="server" ActiveTabIndex="0" Width="100%">
                    <ajaxToolkit:TabPanel ID="TabPanel1" runat="server" HeaderText="Dados Bancários">
                        <ContentTemplate>
                            <div class="bordagrid">
                                <asp:GridView ID="DG" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                    ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="DG_SelectedIndexChanged">
                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <EditRowStyle BackColor="#999999" />
                                    <Columns>
                                        <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                                        <asp:BoundField DataField="CodigoBanco" HeaderText="C&#243;d.">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="NomeBanco" HeaderText="Banco">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="CodigoAgencia" HeaderText="Ag&#234;ncia">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="DigitoAgencia" HeaderText="Dig.">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="ContaCorrente" HeaderText="C.C.">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="DigitoConta" HeaderText="Dig.">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="TipoConta" HeaderText="Tipo">
                                            <HeaderStyle HorizontalAlign="Center" />
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Praca" HeaderText="Pra&#231;a">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Cidade" HeaderText="Cidade">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Estado" HeaderText="Estado">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                    </Columns>
                                </asp:GridView>
                            </div>
                            <div class="row">
                                <div class="painelright">
                                    <asp:Button ID="Button1" runat="server" CssClass="botao" Text="Cadastrar" OnClick="Button1_Click" />
                                    <asp:Button ID="btnFechar" runat="server" CssClass="botao" Text="Fechar" OnClick="btnFechar_Click" />
                                </div>
                            </div>
                        </ContentTemplate>
                    </ajaxToolkit:TabPanel>
                    <ajaxToolkit:TabPanel ID="TabPanel2" runat="server" HeaderText="Cadastro">
                        <ContentTemplate>
                            <div class="row">
                                <div class="collbluc">
                                    Banco:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlBanco" runat="server" Width="400px" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbluc">
                                    Agência:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="TxtCodAgencia" runat="server" Width="120px" />
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="TxtAgenciaDigito" runat="server" Width="50px" MaxLength="1" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbluc">
                                    Conta:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="TxtConta" runat="server" Width="120px" />
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="TxtContaDigito" runat="server" Width="50px" MaxLength="2" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbluc">
                                    Tipo:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlTipoConta" runat="server" Width="123px">
                                        <asp:ListItem Value="C">C. Corrente</asp:ListItem>
                                        <asp:ListItem Value="P">C. Popan&#231;a</asp:ListItem>
                                    </asp:DropDownList>
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbluc">
                                    Praça:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="TxtPraca" runat="server" Width="390px" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbluc">
                                    UF:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlEstado" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlEstado_SelectedIndexChanged"
                                        Width="400px" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbluc">
                                    Cidade:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlCidade" runat="server" Width="400px" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="painelright">
                                    <asp:Button ID="BtnSalvar" runat="server" CssClass="botao" Text="Salvar" OnClick="Button2_Click" />
                                    <asp:Button ID="btnClose" runat="server" CssClass="botao" Text="Fechar" OnClick="btnClose_Click" />
                                </div>
                            </div>
                        </ContentTemplate>
                    </ajaxToolkit:TabPanel>
                </ajaxToolkit:TabContainer>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
