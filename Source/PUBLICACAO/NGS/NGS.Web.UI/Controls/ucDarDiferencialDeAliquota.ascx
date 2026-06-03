<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucDarDiferencialDeAliquota.ascx.vb"
    Inherits="NGS.Web.UI.ucDarDiferencialDeAliquota" %>
<div id="divDarDiferencialDeAliquota" class="uc" title="DAR - Diferencial de Alíquota"
    style="display: none;">
    <asp:UpdatePanel ID="updConsultaViasDeTransportes" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="menu_acoes">
                <div class="acoes" runat="server">
                    <ul>
                        <li class="iconNovo" runat="server">
                            <asp:LinkButton ID="lnkNovo" Text="Gravar" runat="server" />
                        </li>
                        <li class="iconAtualizar" runat="server">
                            <asp:LinkButton ID="lnkAtualizar" Text="Atualizar" runat="server" />
                        </li>
                        <li class="iconExcluir" runat="server">
                            <asp:LinkButton ID="lnkExcluir" Text="Excluir" runat="server" OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" />
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li class="iconRelatorio" runat="server">
                            <asp:LinkButton ID="lnkRelatorio" Text="Relatório" runat="server" />
                        </li>
                        <li class="iconSair" runat="server">
                            <asp:LinkButton ID="lnkFechar" Text="Fechar" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:Label ID="lblEmpresa" runat="server" Width="100%" />
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Cód. Inf. Adic:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlInformacoesAdicionais" runat="server" data-ToolTip="default"
                        ToolTip="Registro E115: Informações Adicionais da Apuração – Valores Declaratórios."
                        Width="634px" />
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Dar:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDar" runat="server" Width="122px" MaxLength="17" Style="text-align: right;" />
                </div>
                <div class="collbluc">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoCliente" runat="server"></asp:HiddenField>
                    <asp:TextBox ID="txtClientes" runat="server" Width="333px" Enabled="False" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnCliente" runat="server" CssClass="btn" UseSubmitBehavior="False"
                        Text=">" />
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Data Ref:
                </div>
                <div class="coltxt" style="width: 132px;">
                    <asp:TextBox ID="txtDataRef" runat="server" Width="100px" CssClass="calendario" />
                </div>
                <div class="collbluc">
                    Data Emissão:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtData" runat="server" Width="100px" CssClass="calendario" />
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Nota:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNota" runat="server" Width="122px" Style="text-align: right;" />
                </div>
                <div class="collbluc">
                    Série:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtSerie" runat="server" Width="122px" MaxLength="3" />
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Código da Receita:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCodigoReceita" runat="server" Width="122px" MaxLength="9" Style="text-align: right;" />
                </div>
                <div class="collbluc">
                    Valor:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtValor" runat="server" Width="122px" Style="text-align: right;"
                        CssClass="txtDecimal" />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="GrdDarDiferencialDeAliquota" runat="server" AutoGenerateColumns="False"
                    CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                    <EditRowStyle BackColor="#999999" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <Columns>
                        <asp:CommandField SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button">
                            <HeaderStyle></HeaderStyle>
                            <ItemStyle Width="25px"></ItemStyle>
                        </asp:CommandField>
                        <asp:BoundField DataField="Dar" HeaderText="Dar">
                            <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="DataReferencia" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Data Referência">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Data" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Data Emissão">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Nome" HeaderText="Cliente">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Endereco" HeaderText="Endereço">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Nota" HeaderText="Nota">
                            <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Serie" HeaderText="Série ">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="CodigoReceita" HeaderText="Cód. da Receita">
                            <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Valor" DataFormatString="{0:N2}" HeaderText="Valor">
                            <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="CodigoSpedInfAdicionaisDeApuracao" HeaderText="Cód. Inf. Adic.">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
