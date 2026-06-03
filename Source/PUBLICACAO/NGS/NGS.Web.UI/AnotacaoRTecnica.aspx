<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="AnotacaoRTecnica.aspx.vb" Inherits="NGS.Web.UI.AnotacaoRTecnica" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngArt" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlArt" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Anotação de Responsabilidade Técnica
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconNovo" ID="lnkNovo" Text="Gravar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAtualizar" ID="lnkAtualizar" Text="Atualizar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconExcluir" ID="lnkExcluir" Text="Excluir" runat="server"
                                OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconLimpar" ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconRelatorio" ID="lnkRelatorio" Text="Relatório" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Codigo:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCodigo" runat="server" MaxLength="20" data-ToolTip="default" ToolTip="Código de cadastro da responsabilidade técnica." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoEmpresa" runat="server" />
                    <asp:TextBox ID="txtEmpresa" runat="server" Width="500px" Enabled="False" data-ToolTip="default" ToolTip="" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnEmpresa" CssClass="btn" OnClick="btnEmpresa_Click" runat="server"
                        UseSubmitBehavior="False" Text=">" data-ToolTip="default" ToolTip="Empresa para negociação/consulta." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Agrônomo:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoAgronomo" runat="server" />
                    <asp:TextBox ID="txtAgronomo" runat="server" Width="500px" Enabled="False" data-ToolTip="default" ToolTip="" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnAgronomo" CssClass="btn" OnClick="btnAgronomo_Click" runat="server"
                        UseSubmitBehavior="False" Text=">" data-ToolTip="default" ToolTip="Nome do agrônomo responsável técnico." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Inicial:
                </div>
                <div class="coltxt" style="width: 130px;">
                    <asp:TextBox ID="txtArtInicial" runat="server" data-ToolTip="default" ToolTip="Número inicial da quantidade de formulários." />
                </div>
                <div class="collbl">
                    Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtArtFinal" runat="server" data-ToolTip="default" ToolTip="Número final da quantidade de formulários." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Validade:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtValidade" CssClass="calendario" runat="server" data-ToolTip="default" ToolTip="Data de vencimento dos formulários agronômicos." />
                </div>
                <div class="collbl">
                    Ativa:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="radSim" runat="server" Text="Sim" Checked="true" ValidationGroup="art"
                        GroupName="art" data-ToolTip="default" ToolTip="Sinalizador do status do formulário." />
                    <asp:RadioButton ID="radNao" runat="server" Text="Não" ValidationGroup="art" GroupName="art"
                        data-ToolTip="default" ToolTip="Sinalizador do status do formulário." />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="gridArt" runat="server" AutoGenerateColumns="False" ForeColor="#333333"
                    GridLines="None" Width="100%" CellPadding="4" OnSelectedIndexChanged="gridArt_SelectedIndexChanged">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:CommandField ButtonType="Button" SelectText="&gt;" ShowSelectButton="True" />
                        <asp:BoundField DataField="CodigoArt" HeaderText="C&#243;digo" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="CodigoEmpresa" HeaderText="Empresa" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="NomeAgronomo" HeaderText="Agr&#244;nomo" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="ARTInicial" HeaderText="RA Inicial" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField DataField="ARTFinal" HeaderText="RA Final" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField DataField="ARTValidade" HeaderText="Validade" DataFormatString="{0:dd/MM/yyyy}">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Status" HeaderText="Status">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaEmpresas ID="ucConsultaEmpresas" runat="server" />
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
