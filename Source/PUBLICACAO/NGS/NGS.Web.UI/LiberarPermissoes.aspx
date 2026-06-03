<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="LiberarPermissoes.aspx.vb" Inherits="NGS.Web.UI.LiberarPermissoes" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrm" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updLiberarPermissoes" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="hdfGrupoId" runat="server" />
            <asp:HiddenField ID="hdfProcessoId" runat="server" />
            <table style="width: 100%;">
                <tr>
                    <td colspan="3">
                        <div class="titulodiv">
                            Liberar Permissões
                        </div>
                    </td>
                </tr>
                <tr>
                    <td colspan="3">
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAjuda" ID="lnkAjuda" Text="Ajuda" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                    </td>
                </tr>
                <tr>
                    <td class="primario" style="width: 33%; height: 10px; text-align: center;">
                        <div class="subtitulodiv">
                            Grupos
                        </div>

                    </td>
                    <td class="primario" style="width: 34%; height: 10px; text-align: center;">
                        <div class="subtitulodiv">
                            Processos
                        </div>
                    </td>
                    <td class="primario" style="width: 33%; height: 10px; text-align: center;">
                        <div class="subtitulodiv">
                            Permissões
                        </div>
                    </td>
                </tr>
                <tr>
                    <td style="padding: 0px 10px 10px 2px;">
                        <asp:UpdatePanel ID="updpnlDivGrupos" runat="server" UpdateMode="Conditional">
                            <ContentTemplate>
                                <div style="overflow: auto; height: 250px; width: 100%; border-color: #CDDBEB; border-style: solid;">
                                    <asp:ListView ID="lstGrupos" runat="server" DataKeyNames="Grupo_Id" OnItemCommand="lstGrupos_OnItemCommand">
                                        <LayoutTemplate>
                                            <table runat="server" width="100%">
                                                <tr runat="server" id="itemPlaceholder">
                                                </tr>
                                            </table>
                                        </LayoutTemplate>
                                        <ItemTemplate>
                                            <tr runat="server">
                                                <td runat="server">
                                                    <asp:UpdatePanel ID="updGrupos" runat="server" UpdateMode="Conditional" ChildrenAsTriggers="false">
                                                        <ContentTemplate>
                                                            <asp:LinkButton ID="lnkGrupos" runat="server" ClientIDMode="AutoID" Style="border: 0;"
                                                                CommandName="AddToList" CommandArgument='<%#Eval("Grupo_Id") %>'>
                                                                <img height="20px" width="20px" src="Images/user_group.ico" data-ToolTip="default" title="Grupos" alt="Grupos" />
                                                            </asp:LinkButton>
                                                        </ContentTemplate>
                                                        <Triggers>
                                                            <asp:AsyncPostBackTrigger ControlID="lnkGrupos" EventName="Click" />
                                                        </Triggers>
                                                    </asp:UpdatePanel>
                                                </td>
                                                <td runat="server">
                                                    <asp:Label ID="lblDescricaoGrupo" runat="server" Text='<%#Eval("Grupo_Id") %>' />
                                                </td>
                                            </tr>
                                        </ItemTemplate>
                                        <AlternatingItemTemplate>
                                            <tr style="background-color: #EFEFEF">
                                                <td runat="server">
                                                    <asp:UpdatePanel ID="updGrupos" runat="server" UpdateMode="Conditional" ChildrenAsTriggers="false">
                                                        <ContentTemplate>
                                                            <asp:LinkButton ID="lnkGrupos1" runat="server" ClientIDMode="AutoID" Style="border: 0;"
                                                                CommandName="AddToList" CommandArgument='<%#Eval("Grupo_Id") %>'>
                                                                <img height="20px" width="20px" src="Images/user_group.ico" data-ToolTip="default" title="Grupos" alt="Grupos" />
                                                            </asp:LinkButton>
                                                        </ContentTemplate>
                                                        <Triggers>
                                                            <asp:AsyncPostBackTrigger ControlID="lnkGrupos1" EventName="Click" />
                                                        </Triggers>
                                                    </asp:UpdatePanel>
                                                </td>
                                                <td runat="server">
                                                    <asp:Label ID="lblDescricaoGrupo" runat="server" Text='<%#Eval("Grupo_Id") %>' />
                                                </td>
                                            </tr>
                                        </AlternatingItemTemplate>
                                    </asp:ListView>
                                </div>
                            </ContentTemplate>
                        </asp:UpdatePanel>
                    </td>
                    <td style="padding: 0px 10px 10px 0px;">
                        <asp:UpdatePanel ID="updpnlDivProcessos" runat="server" UpdateMode="Conditional">
                            <ContentTemplate>
                                <div style="overflow: auto; height: 250px; width: 100%; border-color: #CDDBEB; border-style: solid;">
                                    <asp:ListView ID="lstProcessos" runat="server" DataKeyNames="Processo_Id" OnItemCommand="lstProcessos_OnItemCommand"
                                        Style="overflow: auto; height: 250px; width: 100%; border-color: #CDDBEB; border-style: solid;">
                                        <LayoutTemplate>
                                            <table runat="server" width="100%">
                                                <tr runat="server" id="itemPlaceholder">
                                                </tr>
                                            </table>
                                        </LayoutTemplate>
                                        <ItemTemplate>
                                            <tr>
                                                <td>
                                                    <asp:UpdatePanel ID="updProcessos" runat="server" UpdateMode="Conditional" ChildrenAsTriggers="false">
                                                        <ContentTemplate>
                                                            <asp:LinkButton ID="lnkProcessos" runat="server" Style="border: 0;" CommandName="AddToList"
                                                                CommandArgument='<%#Eval("Processo_Id") %>'>
                                                            <img height="20px" width="20px" src="Images/processo.ico" data-ToolTip="default" title="Processos" alt="Processos" />
                                                            </asp:LinkButton>
                                                        </ContentTemplate>
                                                        <Triggers>
                                                            <asp:AsyncPostBackTrigger ControlID="lnkProcessos" EventName="Click" />
                                                        </Triggers>
                                                    </asp:UpdatePanel>
                                                </td>
                                                <td>
                                                    <asp:Label ID="lblProcesso" runat="server" Text='<%#Eval("Processo_Id") %>' />
                                                </td>
                                            </tr>
                                        </ItemTemplate>
                                        <AlternatingItemTemplate>
                                            <tr style="background-color: #EFEFEF">
                                                <td>
                                                    <asp:UpdatePanel ID="updProcessos" runat="server" UpdateMode="Conditional" ChildrenAsTriggers="false">
                                                        <ContentTemplate>
                                                            <asp:LinkButton ID="lnkProcessos1" runat="server" Style="border: 0;" CommandName="AddToList"
                                                                CommandArgument='<%#Eval("Processo_Id") %>'>
                                                                <img height="20px" width="20px" src="Images/processo.ico" data-ToolTip="default" title="Processos" alt="Processos" />
                                                            </asp:LinkButton>
                                                        </ContentTemplate>
                                                        <Triggers>
                                                            <asp:AsyncPostBackTrigger ControlID="lnkProcessos1" EventName="Click" />
                                                        </Triggers>
                                                    </asp:UpdatePanel>
                                                </td>
                                                <td>
                                                    <asp:Label ID="lblProcesso" runat="server" Text='<%#Eval("Processo_Id") %>' />
                                                </td>
                                            </tr>
                                        </AlternatingItemTemplate>
                                    </asp:ListView>
                                </div>
                            </ContentTemplate>
                        </asp:UpdatePanel>
                    </td>
                    <td style="padding: 0px 10px 10px 0px;">
                        <asp:UpdatePanel ID="updpnlDivPermissoes" runat="server" UpdateMode="Conditional">
                            <ContentTemplate>
                                <div style="overflow: auto; height: 250px; width: 100%; border-color: #CDDBEB; border-style: solid;">
                                    <asp:ListView runat="server" ID="lstPermissoes">
                                        <LayoutTemplate>
                                            <table runat="server" width="100%">
                                                <tr runat="server" id="itemPlaceholder">
                                                </tr>
                                            </table>
                                        </LayoutTemplate>
                                        <ItemTemplate>
                                            <tr runat="server">
                                                <td runat="server">
                                                    <asp:CheckBox ID="chkPermissoes" Text='<%#Eval("Permissao_Id") %>' runat="server" />
                                                </td>
                                            </tr>
                                        </ItemTemplate>
                                        <AlternatingItemTemplate>
                                            <tr style="background-color: #EFEFEF">
                                                <td runat="server">
                                                    <asp:CheckBox ID="chkPermissoes" Text='<%#Eval("Permissao_Id") %>' runat="server" />
                                                </td>
                                            </tr>
                                        </AlternatingItemTemplate>
                                    </asp:ListView>
                                </div>
                            </ContentTemplate>
                        </asp:UpdatePanel>
                    </td>
                </tr>
                <tr>
                    <td colspan="3" style="text-align: center;">
                        <asp:Button ID="btnAtualizar" runat="server" CssClass="botao" Text="Confirmar"></asp:Button>
                    </td>
                </tr>
                <tr>
                    <td class="primario" style="width: 33%; height: 10px; text-align: center;">
                        <div class="subtitulodiv">
                            Usuários
                        </div>
                    </td>
                    <td class="primario" style="width: 34%; height: 10px; text-align: center;">
                        <div class="subtitulodiv">
                            Grupos participantes
                        </div>
                    </td>
                    <td class="primario" style="width: 33%; height: 10px; text-align: center;">
                        <div class="subtitulodiv">
                            Informações
                        </div>
                    </td>
                </tr>
                <tr>
                    <td style="padding: 0px 10px 10px 0px;">
                        <asp:UpdatePanel ID="updpnlDivUsuarios" runat="server" UpdateMode="Conditional">
                            <ContentTemplate>
                                <div style="overflow: auto; height: 250px; width: 100%; border-color: #CDDBEB; border-style: solid;">
                                    <asp:ListView DataKeyNames="Usuario_id" OnItemCommand="lstUsuarios_OnItemCommand"
                                        runat="server" ID="lstUsuarios">
                                        <LayoutTemplate>
                                            <table runat="server" width="100%">
                                                <tr runat="server" id="itemPlaceholder">
                                                </tr>
                                            </table>
                                        </LayoutTemplate>
                                        <ItemTemplate>
                                            <tr>
                                                <td>
                                                    <asp:UpdatePanel ID="updUsuarios" runat="server" UpdateMode="Conditional" ChildrenAsTriggers="false">
                                                        <ContentTemplate>
                                                            <asp:LinkButton ID="lnkUsuario" runat="server" Style="border: 0;" CommandName="AddToList"
                                                                CommandArgument='<%#Eval("Usuario_id") %>'>
                                                                <img height="20px" width="20px" src="Images/man2.png" data-ToolTip="default" title="Usuários" alt="Usuários" />
                                                            </asp:LinkButton>
                                                        </ContentTemplate>
                                                        <Triggers>
                                                            <asp:AsyncPostBackTrigger ControlID="lnkUsuario" EventName="Click" />
                                                        </Triggers>
                                                    </asp:UpdatePanel>
                                                </td>
                                                <td>
                                                    <asp:Label ID="lblUsuario" runat="server" Text='<%#Eval("Usuario_id") %>' />
                                                </td>
                                            </tr>
                                        </ItemTemplate>
                                        <AlternatingItemTemplate>
                                            <tr style="background-color: #EFEFEF">
                                                <td>
                                                    <asp:UpdatePanel ID="updUsuarios" runat="server" UpdateMode="Conditional" ChildrenAsTriggers="false">
                                                        <ContentTemplate>
                                                            <asp:LinkButton ID="lnkUsuario1" runat="server" Style="border: 0;" CommandName="AddToList"
                                                                CommandArgument='<%#Eval("Usuario_id") %>'>
                                                                <img height="20px" width="20px" src="Images/man2.png" data-ToolTip="default" title="Usuários" alt="Usuários" />
                                                            </asp:LinkButton>
                                                        </ContentTemplate>
                                                        <Triggers>
                                                            <asp:AsyncPostBackTrigger ControlID="lnkUsuario1" EventName="Click" />
                                                        </Triggers>
                                                    </asp:UpdatePanel>
                                                </td>
                                                <td>
                                                    <asp:Label ID="lblUsuario" runat="server" Text='<%#Eval("Usuario_id") %>' />
                                                </td>
                                            </tr>
                                        </AlternatingItemTemplate>
                                    </asp:ListView>
                                </div>
                            </ContentTemplate>
                        </asp:UpdatePanel>
                    </td>
                    <td style="padding: 0px 10px 10px 0px;">
                        <asp:UpdatePanel ID="updpnlParticipantes" runat="server" UpdateMode="Conditional">
                            <ContentTemplate>
                                <div style="overflow: auto; height: 250px; width: 100%; border-color: #CDDBEB; border-style: solid;">
                                    <asp:ListView ID="lstParticipantes" runat="server" DataKeyNames="Grupo_Id">
                                        <LayoutTemplate>
                                            <table runat="server" width="100%">
                                                <tr runat="server" id="itemPlaceholder">
                                                </tr>
                                            </table>
                                        </LayoutTemplate>
                                        <ItemTemplate>
                                            <tr runat="server">
                                                <td runat="server">
                                                    <asp:UpdatePanel ID="updParticipantes" runat="server" UpdateMode="Conditional" ChildrenAsTriggers="false">
                                                        <ContentTemplate>
                                                            <asp:LinkButton ID="lnkParticipantes" runat="server" Style="border: 0;" CommandName="AddToList"
                                                                CommandArgument='<%#Eval("Grupo_Id") %>'>
                                                                <img height="20px" width="20px" src="Images/user_group.ico" data-ToolTip="default" title="Participantes" alt="Participantes" />
                                                            </asp:LinkButton>
                                                        </ContentTemplate>
                                                        <Triggers>
                                                            <asp:AsyncPostBackTrigger ControlID="lnkParticipantes" EventName="Click" />
                                                        </Triggers>
                                                    </asp:UpdatePanel>
                                                </td>
                                                <td id="Td8" runat="server">
                                                    <asp:Label ID="lblDescGrupo" runat="server" Text='<%#Eval("Grupo_Id") %>' />
                                                </td>
                                            </tr>
                                        </ItemTemplate>
                                        <AlternatingItemTemplate>
                                            <tr style="background-color: #EFEFEF">
                                                <td runat="server">
                                                    <asp:UpdatePanel ID="updParticipantes" runat="server" UpdateMode="Conditional" ChildrenAsTriggers="false">
                                                        <ContentTemplate>
                                                            <asp:LinkButton ID="lnkParticipantes1" runat="server" Style="border: 0;" CommandName="AddToList"
                                                                CommandArgument='<%#Eval("Grupo_Id") %>'>
                                                                <img height="20px" width="20px" src="Images/user_group.ico" data-ToolTip="default" title="Participantes" alt="Participantes" />
                                                            </asp:LinkButton>
                                                        </ContentTemplate>
                                                        <Triggers>
                                                            <asp:AsyncPostBackTrigger ControlID="lnkParticipantes1" EventName="Click" />
                                                        </Triggers>
                                                    </asp:UpdatePanel>
                                                </td>
                                                <td runat="server">
                                                    <asp:Label ID="lblDescGrupo" runat="server" Text='<%#Eval("Grupo_Id") %>' />
                                                </td>
                                            </tr>
                                        </AlternatingItemTemplate>
                                    </asp:ListView>
                                </div>
                            </ContentTemplate>
                        </asp:UpdatePanel>
                    </td>
                    <td style="padding: 0px 10px 10px 0px;" valign="top">
                        <div style="overflow: auto; height: 250px; width: 100%; border-color: #CDDBEB; border-style: solid;">
                            <table style="width: 100%; height: 95px" border="0">
                                <tr>
                                    <td class="fonteazulbold" style="width: 142px" align="center">Grupo<br />
                                        selecionado:
                                    </td>
                                    <td>
                                        <asp:TextBox runat="server" ID="txtGrupo" Style="width: 128px; color: green" />
                                    </td>
                                </tr>
                                <tr>
                                    <td class="fonteazulbold" style="width: 142px" align="center">Processo selecionado:
                                    </td>
                                    <td>
                                        <asp:TextBox runat="server" ID="txtProcesso" Style="width: 128px; color: green" />
                                    </td>
                                </tr>
                                <tr>
                                    <td class="fonteazulbold" style="width: 142px" align="center">Usuário selecionado:
                                    </td>
                                    <td>
                                        <asp:TextBox runat="server" ID="txtUsuario" Style="width: 128px; color: green" />
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="2" style="text-align: center;">
                                        <asp:Button ID="cmdAjuda" runat="server" Text="Ajuda" CssClass="botao" />
                                        <asp:Button ID="btnSair" runat="server" Text="Sair" CssClass="botao" OnClick="btnSair_Click" />
                                    </td>
                                </tr>
                            </table>
                        </div>
                    </td>
                </tr>
            </table>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
