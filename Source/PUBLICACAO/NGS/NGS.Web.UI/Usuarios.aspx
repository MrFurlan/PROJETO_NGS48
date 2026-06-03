<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="Usuarios.aspx.vb" Inherits="NGS.Web.UI.Usuarios" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        #meioconteudo {
            width: 1100px !important;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngUsuarios" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlUsuarios" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <asp:HiddenField ID="txtCodigoCliente" runat="server" />
            <div class="titulodiv">
                Usuários
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconNovo" runat="server">
                            <asp:LinkButton ID="lnkNovo" Text="Gravar" runat="server" />
                        </li>
                        <li class="iconAtualizar" runat="server">
                            <asp:LinkButton ID="lnkAtualizar" Text="Atualizar" runat="server" />
                        </li>
<%--                        <li class="iconExcluir" runat="server">
                            <asp:LinkButton ID="lnkExcluir" Text="Excluir" runat="server" OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" />
                        </li>--%>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li class="iconRelatorio" runat="server">
                            <asp:LinkButton ID="lnkRelatorio" Text="Relatório" runat="server" />
                        </li>
                        <li class="iconAjuda" runat="server">
                            <asp:LinkButton ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Usuários:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtUsuario" runat="server" Width="160px" MaxLength="50" data-ToolTip="default" ToolTip="Código de cadastro do usuário." />
                    <asp:CheckBox ID="chkAtivo" runat="server" Text="Usuário Ativo" />
                    <asp:CheckBox ID="ChkImprimirDanfe" runat="server" Text="Imprimir Danfe" />
                    <asp:CheckBox ID="chkLiberaEmpresa" runat="server" Text="Liberar Empresa" />
                    <asp:CheckBox ID="chkTrocaEmpresa" runat="server" Text="Trocar Empresa" />
                    <asp:CheckBox ID="chkLiberaJanela" runat="server" Text="Liberar Janela" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Nome Completo:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNomeCompleto" runat="server" Width="645px" MaxLength="50" data-ToolTip="default" ToolTip="Nome Completo do usuário." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Banco de Dados:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlBancoDeDados" runat="server" Width="650px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUnidadeDeNegocio" runat="server" Width="650px" OnSelectedIndexChanged="ddlUnidadeDeNegocio_SelectedIndexChanged"
                        AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="650px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Menu de Acesso:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlMenuDeAcesso" runat="server" Width="650px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Email:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtEmail" runat="server" Width="645px" MaxLength="50" data-ToolTip="default" ToolTip="e-mail do usuário." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente Vinculado:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCliente" runat="server" Width="550px" Enabled="False" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnConsultarCliente" CssClass="btn" runat="server" UseSubmitBehavior="False" Text=">"
                        data-ToolTip="default" ToolTip="Selecionar caso haja cliente vinculado ao usuário." />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="gridUsuarios" runat="server" AutoGenerateColumns="False" CellPadding="3"
                    ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="gridUsuarios_SelectedIndexChanged">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:CommandField SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button">
                            <ItemStyle Width="25px"></ItemStyle>
                        </asp:CommandField>
                        <asp:BoundField DataField="Usuario" HeaderText="Usu&#225;rio" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left" Width="100px"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="NomeCompleto" HeaderText="Nome Completo" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left" Width="200px"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="BancoDeDados" HeaderText="Banco De Dados" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left" Width="100px"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="MenuDeAcesso" HeaderText="Menu De Acesso" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left" Width="100px"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Email" HeaderText="Email" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="AcessoUnidade" HeaderText="Un. Neg&#243;cio" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="AcessoEmpresa" HeaderText="Empresa" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="AcessoEndEmpresa" HeaderText="End" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Cliente" HeaderText="Cliente" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="EndCliente" HeaderText="Endereço" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="NomeCliente" HeaderText="Nome do Cliente" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:TemplateField HeaderText="Ativo">
                            <ItemTemplate>
                                <asp:CheckBox ID="chkAtivogrid" runat="server" Checked='<%# Bind("Ativo") %>' Enabled="false" />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Imprimir">
                            <ItemTemplate>
                                <asp:CheckBox ID="gridchkImprimir" runat="server" Checked='<%# Bind("ImprimirDanfe") %>' Enabled="false" />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Liberar">
                            <ItemTemplate>
                                <asp:CheckBox ID="gridchkLiberaEmpresa" runat="server" Checked='<%# Bind("LiberaEmpresa") %>' Enabled="false" />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Trocar">
                            <ItemTemplate>
                                <asp:CheckBox ID="gridchkTrocaEmpresa" runat="server" Checked='<%# Bind("TrocaEmpresa") %>' Enabled="false" />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Janela">
                            <ItemTemplate>
                                <asp:CheckBox ID="gridchkLiberaJanela" runat="server" Checked='<%# Bind("LiberaJanela") %>' Enabled="false" />
                            </ItemTemplate>
                        </asp:TemplateField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
