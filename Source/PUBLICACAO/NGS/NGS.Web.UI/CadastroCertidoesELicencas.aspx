<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master" CodeBehind="CadastroCertidoesELicencas.aspx.vb" Inherits="NGS.Web.UI.CadastroCertidoesELicencas" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngCadastroCertidoesELicencas" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlCadastroCertidoesELicencas" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Cadastro De Certidões e Licenças
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconNovo" ID="lnkNovo" runat="server" Text="Gravar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAtualizar" ID="lnkAtualizar" runat="server" Text="Atualizar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconExcluir" ID="lnkExcluir" runat="server" Text="Excluir"
                                OnClientClick="return msgconfirm(this);" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconConsultar" ID="lnkConsultar" runat="server" Text="Consultar" />
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
                    Tipo:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlTipo" runat="server" Width="200px" />
                </div>
                <div class="coltxt">
                    <asp:LinkButton ID="lnkAddTipo" runat="server" Style="text-align: center;" data-ToolTip="default"
                        ToolTip="Incluir Novo Tipo De Certidões e Licenças">
                        <img alt="Novo" src="Images/detalhes.png" style="vertical-align:middle; height:20px;"/>
                    </asp:LinkButton>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="640px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data Emissão:
                </div>
                <div class="coltxt" style="width: 210px">
                    <asp:TextBox ID="txtDataEm" runat="server" Width="100px" CssClass="calendario" />
                </div>
                <div class="collbl">
                    Data Vencimento:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataVenc" runat="server" Width="100px" CssClass="calendario" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    E-mail:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtEmail" runat="server" Width="200px" />
                </div>
                <div class="collbl">
                    Aviso Dias:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtAviso" runat="server" Width="100px" CssClass="txtNumerico" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Observação:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtObservacao" runat="server" Width="640px" Height="50px" TextMode="MultiLine" />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="GridCadastroCertidaoELicenca" runat="server" Width="100%" AutoGenerateColumns="False"
                    CellPadding="4" ForeColor="#333333" GridLines="None" OnSelectedIndexChanged="GridCadastroCertidaoELicenca_SelectedIndexChanged">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                        <asp:BoundField DataField="DescricaoTipo" HeaderText="Tipo">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="DataEmissao" HeaderText="Data de Emissao" DataFormatString="{0:dd/MM/yyyy}">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="DataVencimento" HeaderText="Data de Vencimento" DataFormatString="{0:dd/MM/yyyy}">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Email" HeaderText="Email">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="AvisoDias" HeaderText="Avido de Dias">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Observacao" HeaderText="Observacao">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:CadastrarTipoDeCertidao ID="ucCadastrarTipoDeCertidao" runat="server" />
</asp:Content>
