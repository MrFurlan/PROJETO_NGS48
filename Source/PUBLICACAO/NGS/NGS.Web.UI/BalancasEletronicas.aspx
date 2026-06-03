<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="BalancasEletronicas.aspx.vb" Inherits="NGS.Web.UI.BalancasEletronicas" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngBalancaEletronica" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlBalancaEletronica" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Balança Eletrônica
            </div>
            <ajaxToolkit:TabContainer ID="TabContainer1" runat="server" ActiveTabIndex="0" Width="100%">
                <ajaxToolkit:TabPanel ID="TabPanel1" runat="server">
                    <HeaderTemplate>
                        Balanças
                    </HeaderTemplate>
                    <ContentTemplate>
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
                                        <asp:LinkButton class="iconAjuda" ID="lnkAjuda" Text="Ajuda" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="painelleft">
                                <div class="row">
                                    <div class="collbl">
                                        Nome:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtNomeBalanca" runat="server" data-ToolTip="default" ToolTip="Nome da Balança." />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        Eletrônica:
                                    </div>
                                    <div class="coltxt">
                                        <asp:RadioButton ID="radESim" runat="server" GroupName="balanca" Text="Sim" data-ToolTip="default"
                                            ToolTip="Informar se a balança é eletrônica." />
                                        <asp:RadioButton ID="radENao" runat="server" GroupName="balanca" Text="Não" data-ToolTip="default"
                                            ToolTip="Informar se a balança é eletrônica." />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        Porta:
                                    </div>
                                    <div class="coltxt">
                                        <asp:DropDownList ID="ddlPorta" runat="server" Width="110px">
                                            <asp:ListItem>COM1</asp:ListItem>
                                            <asp:ListItem>COM2</asp:ListItem>
                                            <asp:ListItem>COM3</asp:ListItem>
                                            <asp:ListItem>COM4</asp:ListItem>
                                        </asp:DropDownList>
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        DataBits:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtDataBits" runat="server" />
                                    </div>
                                </div>
                            </div>
                            <div class="painelleft">
                                <div class="row">
                                    <div class="collbl">
                                        IP:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtIPBalanca" runat="server" data-ToolTip="default" ToolTip="Identificação da balança." />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        ReadBuffer:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtReadBuffer" runat="server" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        BaudRate:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtBaudRate" runat="server" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        StopBits:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtStopBits" runat="server" />
                                    </div>
                                </div>
                            </div>
                            <div class="painelleft">
                                <div class="row">
                                    <div class="collbl">
                                        Tipo:
                                    </div>
                                    <div class="coltxt">
                                        <asp:DropDownList ID="ddlModelo" runat="server" Width="132px">
                                            <asp:ListItem Value="SAT">SATURNO</asp:ListItem>
                                            <asp:ListItem>EPM</asp:ListItem>
                                            <asp:ListItem Value="TOL">TOLEDO</asp:ListItem>
                                            <asp:ListItem Value="CAP">CAPITOL</asp:ListItem>
                                            <asp:ListItem>WEI</asp:ListItem>
                                            <asp:ListItem Value="AEP">AEP DO BRASIL</asp:ListItem>
                                        </asp:DropDownList>
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        ReceivedBytes:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtReceiveBytes" runat="server" Width="122px" data-ToolTip="default"
                                            ToolTip="" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        Paridade:
                                    </div>
                                    <div class="coltxt">
                                        <asp:DropDownList ID="ddlParidade" runat="server" Width="132px">
                                            <asp:ListItem Value="0">Nenhuma</asp:ListItem>
                                            <asp:ListItem Value="1">Impar</asp:ListItem>
                                            <asp:ListItem Value="2">Par</asp:ListItem>
                                        </asp:DropDownList>
                                    </div>
                                </div>
                            </div>
                        </div>
                        <div class="row">
                            <div class="bordagrid">
                                <asp:GridView ID="gridBalancaEletronica" runat="server" AutoGenerateColumns="False"
                                    CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="gridBalancaEletronica_SelectedIndexChanged">
                                    <EditRowStyle BackColor="#999999" />
                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <Columns>
                                        <asp:CommandField ButtonType="Button" SelectText="&gt;" ShowSelectButton="True" />
                                        <asp:BoundField DataField="BalancaServidor" HeaderText="Balan&#231;a">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="BalancaIp" HeaderText="Ip Balan&#231;a">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="BalancaTipo" HeaderText="Tipo">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Eletronica" HeaderText="Eletr&#244;nica">
                                            <HeaderStyle HorizontalAlign="Center" />
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="ReadBuffer" HeaderText="Read Buffer">
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="ReceivedBytes" HeaderText="Received Bytes">
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="PortName" HeaderText="Porta">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="BaudRate" HeaderText="Baud Rate">
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Parity" HeaderText="Paridade">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="DataBits" HeaderText="DataBits">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="StopBits" HeaderText="StopBits">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                    </Columns>
                                </asp:GridView>
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel2" runat="server">
                    <HeaderTemplate>
                        Usuários
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li class="iconNovo" runat="server">
                                        <asp:LinkButton ID="lnkNovoU" Text="Gravar" runat="server" />
                                    </li>
                                    <li class="iconExcluir" runat="server">
                                        <asp:LinkButton ID="lnkExcluirU" runat="server" Text="Excluir" OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" />
                                    </li>
                                    <li class="iconLimpar" runat="server">
                                        <asp:LinkButton ID="lnkLimparU" Text="Limpar" runat="server" />
                                    </li>
                                    <li class="iconAjuda" runat="server">
                                        <asp:LinkButton ID="lnkAjudaU" Text="Ajuda" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Usuário:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlUsuario" runat="server" Width="300px" data-ToolTip="default"
                                    ToolTip="Nome dos usuários." />
                            </div>
                            <div class="collbl">
                                IP:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtIpUsuario" runat="server" Enabled="False" data-ToolTip="default"
                                    ToolTip="Identificação da balança." />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="gridBalancaXUsuarios" runat="server" AutoGenerateColumns="False"
                                CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="gridBalancaXUsuarios_SelectedIndexChanged">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <EditRowStyle BackColor="#999999" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <Columns>
                                    <asp:CommandField ButtonType="Button" SelectText="&gt;" ShowSelectButton="True" />
                                    <asp:BoundField DataField="CodigoBalancaServidor" HeaderText="Balan&#231;a">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="BalancaUsuarioIp" HeaderText="IP Usu&#225;rio">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="NomeUsuario" HeaderText="Usu&#225;rio">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
