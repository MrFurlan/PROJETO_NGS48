<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="CessaoDeCredito.aspx.vb" Inherits="NGS.Web.UI.CessaoDeCredito" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngProcuracao" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlProcuracao" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Cessão de Crédito
            </div>
            <ajaxToolkit:TabContainer ID="TabContainer1" runat="server" ActiveTabIndex="0" Width="100%">
                <ajaxToolkit:TabPanel runat="server" HeaderText="TabConsulta" ID="TabConsulta">
                    <HeaderTemplate>
                        Consulta
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconConsultar" ID="lnkConsultar" Text="Consultar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconRelatorio" ID="lnkRelatorio" Text="Relatório" runat="server" />
                                    </li>
                                    <li runat="server" style="font-weight: 700">
                                        <asp:LinkButton class="iconAjuda" ID="lnkAjuda" Text="Ajuda" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>

                        <div class="row">
                            <div class="collbl">
                                Empresa:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlEmpresaC" runat="server" Width="628px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Cedente:
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="HDCedenteC" runat="server" />
                                <asp:TextBox ID="txtCedenteC" runat="server" Width="590px" Enabled="False" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnCendenteC" runat="server" Text=">" CssClass="btn" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Cessionário:
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="HDCessionarioC" runat="server" />
                                <asp:TextBox ID="txtCessionarioC" runat="server" Width="590px" Enabled="False" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnCessionarioC" runat="server" Text=">" CssClass="btn" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Safra:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlSafraC" runat="server" Width="361px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Cessão de Crédito:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtCessaoDeCreditoC" runat="server" data-tooltip="default" ToolTip="Número da Sessão de Crédito - Se Informado, o sistema irá ignorar qualquer outro parâmetro existente. " />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                <asp:CheckBox ID="chkUsarPeriodo" Text="Usar Periodo:" runat="server" AutoPostBack="True" />

                            </div>
                            <div class="coltxt" id="divPeriodo" runat="server" visible="False">
                                <asp:TextBox ID="txtDataInicialC" CssClass="calendario" runat="server" Width="80px" />
                                &nbsp;a &nbsp;
                                <asp:TextBox ID="txtDataFinalC" CssClass="calendario" runat="server" Width="80px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Consultar:
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="rdPendentes" Text="Pendentes" runat="server" GroupName="x" Checked="True" />
                                <asp:RadioButton ID="rdLiquidados" Text="Liquidados" runat="server" GroupName="x" />
                                <asp:RadioButton ID="rdTodos" Text="Todos" runat="server" GroupName="x" />
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" HeaderText="TabPanel1" ID="TabPanel1">
                    <HeaderTemplate>
                        Cadastro
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconNovo" ID="lnkNovo" Text="Gravar" runat="server" />
                                    </li>
                                    <li runat="server" visible="false">
                                        <asp:LinkButton class="iconAtualizar" ID="lnkAtualizar" Text="Atualizar" runat="server" />
                                    </li>
                                    <li runat="server" visible="false">
                                        <asp:LinkButton class="iconExcluir" ID="lnkExcluir" Text="Excluir" runat="server"
                                            OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconLimpar" ID="lnkLimpar" Text="Limpar" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Empresa:
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="txtCodigoEmpresa" runat="server" />
                                <asp:TextBox ID="txtEmpresa" runat="server" Width="590px" Enabled="False" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnEmpresa" runat="server" Text=">" CssClass="btn" OnClick="btnEmpresa_Click" />
                            </div>

                        </div>
                        <div class="row">
                            <div class="collbl">
                                Cessão de Créd:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtCessaoDeCredito" runat="server" Width="102px" />
                            </div>
                            <div class="collbl">
                                Situação:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlSituacoes" runat="server" Enabled="False" Width="120px">
                                    <asp:ListItem></asp:ListItem>
                                    <asp:ListItem Value="1">Normal</asp:ListItem>
                                    <asp:ListItem Value="2">Cancelado</asp:ListItem>
                                    <asp:ListItem Value="3">Excluído</asp:ListItem>
                                </asp:DropDownList>
                            </div>
                            <div class="collbl">
                                Usuário:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlUsuarios" runat="server" Width="140px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Cedente:
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="txtCodigoCedente" runat="server" />
                                <asp:TextBox ID="txtCedente" runat="server" Width="590px" Enabled="False" />
                            </div>

                            <div class="coltxt">
                                <asp:Button ID="btnCedente" runat="server" Text=">" CssClass="btn" OnClick="btnCedente_Click" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Pedido Cedente:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtPedidoCedente" runat="server" Width="100px" Enabled="False" />
                            </div>

                            <div class="coltxt">
                                <asp:Button ID="btnPedidoCedente" runat="server" Text=">" CssClass="btn" OnClick="btnPedidoCedente_Click" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Cessionário:
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="txtCodigoCessionario" runat="server" />
                                <asp:TextBox ID="txtCessionario" runat="server" Width="590px" Enabled="False" />
                            </div>

                            <div class="coltxt">
                                <asp:Button ID="btnCessionario" runat="server" Text=">" CssClass="btn" OnClick="btnCessionario_Click" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Data Movimento:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtMovimento" CssClass="calendario" runat="server" Width="80px" />
                            </div>

                            <div class="collbl">
                                Documento:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDocumento" runat="server" Width="106px" />
                            </div>

                            <div class="collbl">
                                Quantidade:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtQuantidade" runat="server" Width="106px" />
                            </div>

                        </div>
                        <div class="row">
                            <div class="collbl">
                                Observações:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtObservacoes" runat="server" Height="100px" TextMode="MultiLine"
                                    Width="590px" />
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" HeaderText="TabPanel1" ID="TabPanel2">
                    <HeaderTemplate>
                        Movimentos
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="bordagrid">
                            <asp:GridView ID="grdMovimentos" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:BoundField DataField="Descricao" HeaderText="Descrição" />
                                    <asp:BoundField DataField="Movimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Data" />
                                    <asp:BoundField DataField="Realizado" DataFormatString="{0:N4}" HeaderText="Realizado"
                                        HtmlEncode="False" />
                                </Columns>
                            </asp:GridView>
                        </div>
                        <div class="subtitulodiv">
                            <div class="coltxt right">
                                <asp:Label ID="lblRealizado" runat="server" Text="0.0000" />
                            </div>
                            <div class="coltxt right">
                                Realizado:
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaEmpresas ID="ucConsultaEmpresas" runat="server" />
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaPedidos ID="ucConsultaPedidos" runat="server" />
    <uc:ConsultaProcuracao ID="ucConsultaProcuracao" runat="server" />
</asp:Content>
