<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="Configuracoes.aspx.vb" Inherits="NGS.Web.UI.Configuracoes" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <script type="text/javascript">
        function selectAllNaoVinculados(chkAll) {
            var chk = $('#' + chkAll.id);
            var checked = chk.attr('checked') == "checked";
            $("input[type='checkbox']", "#MainContent_grdNaoVinculados").not("#chkAll").each(function () {
                $(this).attr("checked", checked);
            });
        }

        function selectAllVinculados(chkAll) {
            var chk = $('#' + chkAll.id);
            var checked = chk.attr('checked') == "checked";
            $("input[type='checkbox']", "#MainContent_grdVinculados").not(".chkAlls").each(function () {
                $(this).attr("checked", checked);
            });
        }
    </script>
    <style>
        .collbl
        {
            width: 150px;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngConfiguracoes" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlConfiguracoes" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Configurações
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconNovo" ID="lnkNovo" Text="Salvar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconSair" ID="lnkSair" Text="Sair" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="subtitulodiv">
                E-Mail
            </div>
            <div class="row">
                <div class="collbl">
                    Código:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCodigo" runat="server" Width="150px" CssClass="txtNumerico" Enabled="false" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    E-mail:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtEmail" runat="server" Width="500px" MaxLength="100" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Host:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtHost" runat="server" Width="500px" MaxLength="100" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Senha:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtSenha" runat="server" Type="password" Width="500px" MaxLength="50" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Usuário:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtUsuario" runat="server" Width="500px" MaxLength="100" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Porta:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPorta" runat="server" Width="150px" CssClass="txtNumerico" MaxLength="5" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkSsl" runat="server" Text="SSL?" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkCredenciais" runat="server" Text="Credenciais padrão?" />
                </div>
            </div>
            <div class="subtitulodiv">
                Web Services
            </div>
            <div class="row">
                <div class="collbl">
                    WS Cheque:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCheque" runat="server" Width="500px" MaxLength="100" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    WS Balança:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtBalanca" runat="server" Width="500px" MaxLength="100" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    WS FilesServer:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtFilesServer" runat="server" Width="500px" MaxLength="100" />
                </div>
            </div>
            <div class="subtitulodiv">
                Etapas
            </div>
            <div class="row">
                <div class="collbl">
                    Etapas:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdoCheque" runat="server" Text="Cheque" GroupName="Etapas" Checked="true"
                        AutoPostBack="True" />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdoNFePendencias" runat="server" Text="NF-e Pendências" GroupName="Etapas"
                        AutoPostBack="True" />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdoEstoqueMinimo" runat="server" Text="Estoque Mínimo" GroupName="Etapas"
                        AutoPostBack="True" />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdDFe" runat="server" Text="DFe" GroupName="Etapas"
                        AutoPostBack="True" />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdContabilizarNotas" runat="server" Text="Contabilizar Notas" GroupName="Etapas"
                        AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Email(s):
                </div>
            </div>
            <div class="painelleft" style="width: 48%;">
                <div class="subtitulodiv">
                    Não vinculados
                </div>
                <div class="bordagrid">
                    <asp:GridView ID="grdNaoVinculados" runat="server" AutoGenerateColumns="False" CellPadding="4"
                        DataKeyNames="Usuario_Id" ForeColor="#333333" GridLines="None" Width="100%">
                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                        <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                        <EditRowStyle BackColor="#999999" />
                        <Columns>
                            <asp:TemplateField>
                                <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="5px" />
                                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="5px" />
                                <HeaderTemplate>
                                    <input id="chkAll" onclick="selectAllNaoVinculados(this);" type="checkbox" />
                                </HeaderTemplate>
                                <ItemTemplate>
                                    <asp:CheckBox ID="chkSelecionado" runat="server" />
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:BoundField DataField="Usuario_Id" HeaderText="Usuário">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" Width="100px" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Email" HeaderText="E-mail">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                        </Columns>
                    </asp:GridView>
                </div>
            </div>
            <div class="painelleft" style="width: 4%; text-align: center; margin-top: 200px;">
                <asp:LinkButton ID="lnkDireita" runat="server">
                    <img src="Images/next16.png" alt="direita" />
                </asp:LinkButton>
                <br />
                <asp:LinkButton ID="lnkEsquerda" runat="server">
                    <img src="Images/back.png" alt="esquerda" />
                </asp:LinkButton>
            </div>
            <div class="painelleft" style="width: 48%;">
                <div class="subtitulodiv">
                    Vinculados
                </div>
                <div class="bordagrid">
                    <asp:GridView ID="grdVinculados" runat="server" AutoGenerateColumns="False" CellPadding="4"
                        DataKeyNames="Usuario_Id" ForeColor="#333333" GridLines="None" Width="100%">
                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                        <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                        <EditRowStyle BackColor="#999999" />
                        <Columns>
                            <asp:TemplateField>
                                <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="5px" />
                                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="5px" />
                                <HeaderTemplate>
                                    <input id="chkAlls" onclick="selectAllVinculados(this);" type="checkbox" />
                                </HeaderTemplate>
                                <ItemTemplate>
                                    <asp:CheckBox ID="chkSelecionado" runat="server" />
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:BoundField DataField="Usuario_Id" HeaderText="Usuário">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" Width="100px" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Email" HeaderText="E-mail">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                        </Columns>
                    </asp:GridView>
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
