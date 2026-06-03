<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="UsuariosXProcessos.aspx.vb" Inherits="NGS.Web.UI.UsuariosXProcessos" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <script type="text/javascript">
        function pageLoadUsuariosXProcessos() {
        }
        $(document).ready(function () {
            pageLoadUsuariosXProcessos();
            var prmUsuariosXProcessos = Sys.WebForms.PageRequestManager.getInstance();
            prmUsuariosXProcessos.add_endRequest(pageLoadUsuariosXProcessos);
        });

    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmGruposXUsuarios" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlGruposXUsuarios" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Liberar Permissões
            </div>
            <ajaxToolkit:TabContainer Width="100%" Style="margin-top: 4px;" ID="TabContainer1"
                runat="server" ActiveTabIndex="0">
                <ajaxToolkit:TabPanel ID="TabGrupo" runat="server" HeaderText="TabPanel1">
                    <HeaderTemplate>
                        Por Grupo
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li class="iconNovo" runat="server">
                                        <asp:LinkButton ID="lnkNovoGrupo" runat="server" Text="Gravar" />
                                    </li>
                                    <li class="iconExcluir" runat="server">
                                        <asp:LinkButton ID="lnkExcluirGrupo" runat="server" Text="Excluir"
                                            data-ToolTip="default" ToolTip="Retira o Processo do Grupo e suas permissões."
                                            OnClientClick="if(!confirm('ATENÇÂO\n\nDeseja realmente excluir este Processo?\n\nSe houver permissões as mesmas serão excluídas.')) return false;" />
                                    </li>
                                    <li class="iconLimpar" runat="server">
                                        <asp:LinkButton ID="lnkLimparGrupo" Text="Limpar" runat="server" />
                                    </li>
                                    <li class="iconRelatorio" runat="server">
                                        <asp:LinkButton ID="lnkRelatorioGrupo" Text="Relatório" runat="server" />
                                    </li>
                                    <li class="iconAjuda" runat="server">
                                        <asp:LinkButton ID="lnkAJudaGrupo" Text="Ajuda" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Grupo:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlGrupo" CssClass="multiselect" TabIndex="3" runat="server"
                                    Width="550px" OnSelectedIndexChanged="ddlGrupo_SelectedIndexChanged" AutoPostBack="True" />
                            </div>
                            <div class="coltxt">
                                <asp:LinkButton ID="lnkAddGrupo" runat="server" Style="text-align: center;" data-ToolTip="default"
                                    ToolTip="Incluir Novo Grupo">
                                    <img alt="Novo" src="Images/detalhes.png" style="vertical-align:middle; height:20px;"/>
                                </asp:LinkButton>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Processo:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlProcessoGrupo" TabIndex="4" runat="server" Width="550px" />
                            </div>
                            <div class="coltxt">
                                <asp:LinkButton ID="lnkAddProcessoGrupo" runat="server" Style="text-align: center;"
                                    data-ToolTip="default" ToolTip="Incluir Novo Processo">
                                    <img alt="Novo" src="Images/detalhes.png" style="vertical-align:middle; height:20px;"/>
                                </asp:LinkButton>
                            </div>
                        </div>
                        <div class="row">
                            <div class="painelleft" style="width: 49%; margin-top: 0;">
                                <div class="subtitulodiv">
                                    Processos do Grupo
                                </div>
                                <div class="bordagrid" style="height: 360px;">
                                    <asp:GridView ID="grdProcessoGrupo" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                        ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="grdProcessoGrupo_SelectedIndexChanged"
                                        ShowHeader="False">
                                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                        <EditRowStyle BackColor="#999999" />
                                        <Columns>
                                            <asp:CommandField SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button">
                                                <ItemStyle Width="25px"></ItemStyle>
                                            </asp:CommandField>
                                            <asp:BoundField DataField="Grupo" HeaderText="Grupos">
                                                <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Left" Width="100px"></ItemStyle>
                                            </asp:BoundField>
                                            <asp:BoundField DataField="Processo" HeaderText="Processos">
                                                <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                            </asp:BoundField>
                                        </Columns>
                                    </asp:GridView>
                                </div>
                            </div>
                            <div class="painelright" style="width: 50%; margin-top: 0;">
                                <div class="subtitulodiv">
                                    Permisões do Processo
                                </div>
                                <div class="bordagrid" style="height: 360px;">
                                    <asp:GridView ID="grdPermissoesGrupo" runat="server" AutoGenerateColumns="False"
                                        CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%" ShowHeader="False">
                                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                        <EditRowStyle BackColor="#999999" />
                                        <Columns>
                                            <asp:TemplateField>
                                                <ItemTemplate>
                                                    <asp:CheckBox ID="chkPermitido" runat="server" Checked='<%# Eval("Permitido") %>'
                                                        AutoPostBack="True" OnCheckedChanged="chkPermissaoGrupo_CheckedChanged" />
                                                </ItemTemplate>
                                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center" Width="10px"></ItemStyle>
                                            </asp:TemplateField>
                                            <asp:BoundField DataField="Permissao_Id" HeaderText="Permissões">
                                                <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                            </asp:BoundField>
                                        </Columns>
                                    </asp:GridView>
                                </div>
                            </div>
                        </div>
                        <div class="row">
                            <div class="coltxt" style="width: 100%; margin-top: 3px;">
                                <div class="subtitulodiv">
                                    Usuários Pertencentes
                                </div>
                            </div>
                            <div class="coltxt" style="width: 98%; line-height: 20px;">
                                <div class="bordagrid" style="height: 45px; padding-left: 10px; padding-right: 10px;">
                                    <asp:Label ID="lblUsuariosDoGrupo" runat="server" />
                                </div>
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabUsuario" runat="server" HeaderText="TabPanel1">
                    <HeaderTemplate>
                        Por Usuário
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li class="iconNovo" runat="server">
                                        <asp:LinkButton ID="lnkNovo" Text="Inserir" runat="server" />
                                    </li>
                                    <li class="iconAtualizar" runat="server">
                                        <asp:LinkButton ID="lnkExcluir" runat="server" Text="Excluir"
                                            data-ToolTip="default" ToolTip="Retira o Processo do usuário e suas permissões."
                                            OnClientClick="if(!confirm('ATENÇÂO\n\nDeseja realmente excluir este Processo?\n\nSe houver permissões as mesmas serão excluídas.')) return false;" />
                                    </li>
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
                                Usuário:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlUsuarios" CssClass="multiselect" TabIndex="3" runat="server"
                                    Width="550px" OnSelectedIndexChanged="ddlUsuarios_SelectedIndexChanged" AutoPostBack="True"
                                    data-ToolTip="default" ToolTip="Nome dos usuários." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Processo:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlProcessos" TabIndex="4" runat="server" Width="550px" data-ToolTip="default"
                                    ToolTip="Processos liberados por usuário." />
                            </div>
                            <div class="coltxt">
                                <asp:LinkButton ID="lnkAddProcesso" runat="server" Style="text-align: center;" data-ToolTip="default"
                                    ToolTip="Incluir Novo Pocesso">
                                                                <span><img alt="Novo" src="Images/detalhes.png" style="vertical-align:middle; height:20px;"/></span>
                                </asp:LinkButton>
                            </div>
                        </div>
                        <div class="painelleft" style="width: 463px; margin-top: 0;">
                            <div class="subtitulodiv">
                                Processos do Usuário
                            </div>
                            <div class="bordagrid" style="height: 360px;">
                                <asp:GridView ID="grdUsuarios" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                    ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="grdUsuarios_SelectedIndexChanged">
                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <EditRowStyle BackColor="#999999" />
                                    <Columns>
                                        <asp:CommandField SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button">
                                            <ItemStyle Width="25px"></ItemStyle>
                                        </asp:CommandField>
                                        <asp:BoundField DataField="Usuarios" HeaderText="Usuários">
                                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Left" Width="100px"></ItemStyle>
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Processos" HeaderText="Processos">
                                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                        </asp:BoundField>
                                    </Columns>
                                </asp:GridView>
                            </div>
                        </div>
                        <div class="painelright" style="width: 463px; margin-top: 0;">
                            <div class="subtitulodiv">
                                Permisões do Processo
                            </div>
                            <div class="bordagrid" style="height: 360px;">
                                <asp:GridView ID="grdPermissoes" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                    ForeColor="#333333" GridLines="None" Width="100%">
                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <EditRowStyle BackColor="#999999" />
                                    <Columns>
                                        <asp:TemplateField>
                                            <ItemTemplate>
                                                <asp:CheckBox ID="chkPermitido" runat="server" Checked='<%# Eval("Permitido") %>'
                                                    AutoPostBack="True" OnCheckedChanged="chkPermitido_CheckedChanged" />
                                            </ItemTemplate>
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" Width="10px"></ItemStyle>
                                        </asp:TemplateField>
                                        <asp:BoundField DataField="Permissao_Id" HeaderText="Permissoes">
                                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                        </asp:BoundField>
                                    </Columns>
                                </asp:GridView>
                            </div>
                        </div>
                        <div class="row">
                            <div class="coltxt" style="width: 100%;">
                                <div class="subtitulodiv">
                                    Grupos Participantes
                                </div>
                            </div>
                        </div>
                        <div class="row">
                            <div class="coltxt" style="width: 98%; line-height: 20px;">
                                <div class="bordagrid" style="height: 45px; padding-left: 10px; padding-right: 10px;">
                                    <asp:Label ID="lblGruposDoUsuario" runat="server" />
                                </div>
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabCopiaUsuario" runat="server" HeaderText="TabPanel1">
                    <HeaderTemplate>
                        Copiar Usuário
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li class="iconNovo" runat="server">
                                        <asp:LinkButton ID="lnkCopiaUsuario" Text="Copiar" runat="server" />
                                    </li>
                                    <li class="iconLimpar" runat="server">
                                        <asp:LinkButton ID="lnkCopiaLimpar" Text="Limpar" runat="server" />
                                    </li>
                                    <li class="iconAjuda" runat="server">
                                        <asp:LinkButton ID="lnkCopiaAjuda" Text="Ajuda" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Usuário  Origem:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlUsuarioOrigem" CssClass="multiselect" TabIndex="3" runat="server"
                                    Width="550px" OnSelectedIndexChanged="ddlUsuarioOrigem_SelectedIndexChanged" AutoPostBack="True"
                                    data-ToolTip="default" ToolTip="Usuário Origem." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Usuário copia:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlUsuarioCopia" CssClass="multiselect" TabIndex="3" runat="server"
                                    Width="550px" OnSelectedIndexChanged="ddlUsuarioCopia_SelectedIndexChanged" AutoPostBack="True"
                                    data-ToolTip="default" ToolTip="Usuário Copia." />
                            </div>
                        </div>
                        <div class="painelleft" style="width: 463px; margin-top: 0;">
                            <div class="subtitulodiv">
                                Usuário Origem
                            </div>
                            <div class="bordagrid" style="height: 200px;">
                                <asp:GridView ID="grdOrigem" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                    ForeColor="#333333" GridLines="None" Width="100%">
                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                    <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <EditRowStyle BackColor="#999999" />
                                    <Columns>
                                        <asp:BoundField DataField="Processos">
                                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                        </asp:BoundField>
                                    </Columns>
                                </asp:GridView>
                            </div>
                        </div>
                        <div class="painelright" style="width: 463px; margin-top: 0;">
                            <div class="subtitulodiv">
                                Usuário Copia
                            </div>
                            <div class="bordagrid" style="height: 200px;">
                                <asp:GridView ID="grdCopia" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                    ForeColor="#333333" GridLines="None" Width="100%">
                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                    <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <EditRowStyle BackColor="#999999" />
                                    <Columns>
                                        <asp:BoundField DataField="Processos">
                                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                        </asp:BoundField>
                                    </Columns>
                                </asp:GridView>
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel1" runat="server" HeaderText="TabPanel1">
                    <HeaderTemplate>
                        Consultas
                    </HeaderTemplate>
                    <ContentTemplate>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:CadastrarProcesso ID="ucCadastrarProcesso" runat="server" />
    <uc:CadastrarGrupo ID="ucCadastrarGrupo" runat="server" />
</asp:Content>
