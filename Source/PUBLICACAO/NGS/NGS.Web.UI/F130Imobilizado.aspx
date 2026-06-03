<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="F130Imobilizado.aspx.vb" Inherits="NGS.Web.UI.F130Imobilizado" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <script type="text/javascript">
        $(document).ready(function() {
            $('F130Imobilizado').keypress(function(e) {
                if (e.keyCode == 13) {
                    return false;
                }
            });
        });
    </script>
</asp:Content>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmF130Imobilizado" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlF130Imobilizado" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Registro F130 Imobilizado
            </div>

            <ajaxToolkit:TabContainer ID="TabContainer1" runat="server" ActiveTabIndex="0" Width="100%" Style="margin-top: 4px;">
                <ajaxToolkit:TabPanel ID="TabPanel1" runat="server">
                    <HeaderTemplate>
                        F130 Imobilizado
                    </HeaderTemplate>
                    <ContentTemplate>
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
                                            OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" />
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
                        <div class="row" style="margin-top: 10px;">
                            <div class="collbl">
                                Unidade:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlUnidadeDeNegocio" runat="server" Width="633px" AutoPostBack="True" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Empresa:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlEmpresa" runat="server" Width="633px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Movimento_Id:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtMovimentoId" runat="server" CssClass="calendario" Width="100px"
                                    data-ToolTip="default" ToolTip="Período inicial ddo movimento." />
                            </div>
                            <div class="collbl">
                                Nat_Bc_Cred:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtNatBcCred" runat="server" Width="110px" CssClass="txtNumerico" data-ToolTip="default"
                                    ToolTip="" />
                            </div>
                            <div class="collbl" style="width: 142px;">
                                Situacao:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtSituacao" runat="server" Width="88px" MaxLength="1" Enabled="false" data-ToolTip="default"
                                    ToolTip="" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Ident_Bem_Imob:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlIdentBemImob" runat="server" Width="380px">
                                    <asp:ListItem Value=""> </asp:ListItem>
                                    <asp:ListItem Value="1">1 - Edificações e Benfeitorias</asp:ListItem>
                                    <asp:ListItem Value="3">3 - Instalações</asp:ListItem>
                                    <asp:ListItem Value="4">4 - Máquinas</asp:ListItem>
                                    <asp:ListItem Value="5">5 - Equipamentos</asp:ListItem>
                                    <asp:ListItem Value="6">6 - Veículos</asp:ListItem>
                                    <asp:ListItem Value="9">9 - Outros bens incorporados ao Ativo Imobilizado</asp:ListItem>
                                </asp:DropDownList>
                            </div>
                            <div class="collbl" style="width: 142px;">
                                Parc_Oper_Nao_Bc_Cred:
                            </div>
                            <div class="coltxt" style="width: 126px;">
                                <asp:TextBox ID="txtParcOperNaoBcCred" runat="server" Width="88px" MaxLength="2" data-ToolTip="default"
                                    ToolTip="" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Ind_Orig_Cred:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlIndOrigCred" runat="server" Width="634px">
                                    <asp:ListItem Value="0">0 - Aquisição no Mercado Interno</asp:ListItem>
                                    <asp:ListItem Value="1">1 - Aquisição no Mercado Externo (Importação)</asp:ListItem>
                                </asp:DropDownList>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Ind_Util_Bem_Imob:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlIndUtilBemImob" runat="server" Width="634px">
                                    <asp:ListItem Value="1">1 – Produção de Bens Destinados a Venda</asp:ListItem>
                                    <asp:ListItem Value="2">2 - Prestação de Serviços</asp:ListItem>
                                    <asp:ListItem Value="3">3 – Locação a Terceiros</asp:ListItem>
                                    <asp:ListItem Value="9">9 - Outros</asp:ListItem>
                                </asp:DropDownList>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Ind_Nr_Parc:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlIndNrParc" runat="server" Width="634px">
                                    <asp:ListItem Value="1">1 - Integral (Mês de Aquisição)</asp:ListItem>
                                    <asp:ListItem Value="2">2 – 12 Meses</asp:ListItem>
                                    <asp:ListItem Value="3">3 – 24 Meses</asp:ListItem>
                                    <asp:ListItem Value="4">4 – 48 Meses</asp:ListItem>
                                    <asp:ListItem Value="5">5 – 6 Meses (Embalagens de bebidas frias)</asp:ListItem>
                                    <asp:ListItem Value="9">9 - Outra periodicidade definida em Lei</asp:ListItem>
                                </asp:DropDownList>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Cst_PisCofins:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlCstPisCofins" runat="server" Width="634px" data-ToolTip="default"
                                    ToolTip="" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Vl_Oper_Aquis:
                            </div>
                            <div class="coltxt" style="width: 132px;">
                                <asp:TextBox ID="txtValor" runat="server" Width="115px" data-ToolTip="default" CssClass="txtDecimal"
                                    ToolTip="" />
                            </div>
                            <div class="collbl">
                                Vl_Bc_Cred:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtVlBcCred" runat="server" Width="115px" CssClass="txtDecimal" data-ToolTip="default"
                                    ToolTip="" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Aliq_Pis:
                            </div>
                            <div class="coltxt" style="width: 132px;">
                                <asp:TextBox ID="txtAliqPis" runat="server" Width="115px" CssClass="txtDecimal" data-ToolTip="default"
                                    ToolTip="" />
                            </div>
                            <div class="collbl">
                                Vl_Pis:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtVlPis" runat="server" Width="115px" CssClass="txtDecimal" data-ToolTip="default"
                                    ToolTip="" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Aliq_Cofins:
                            </div>
                            <div class="coltxt" style="width: 132px;">
                                <asp:TextBox ID="txtAliqCofins" runat="server" Width="115px" CssClass="txtDecimal" data-ToolTip="default"
                                    ToolTip="" />
                            </div>
                            <div class="collbl">
                                Vl_Cofins:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtVlCofins" runat="server" Width="113px" CssClass="txtDecimal" data-ToolTip="default"
                                    ToolTip="" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Cod_Cta:
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="hdfConta" runat="server" />
                                <asp:TextBox ID="txtConta" runat="server" Width="594px" Enabled="false" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnConsultaConta" runat="server" Text=">" CssClass="btn" data-ToolTip="default"
                                    ToolTip="Selecionar a conta desejada." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Cod_Ccus:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlCodCcus" runat="server" Width="634px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Desc_Bem_Imob:
                            </div>
                            <div class="coltxt" style="width: 140px;">
                                <asp:TextBox ID="txtDescBemImob" runat="server" TextMode="MultiLine" Width="625px" Height="45px" MaxLength="50" data-ToolTip="default"
                                    ToolTip="" />
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel2" runat="server">
                    <HeaderTemplate>
                        Consulta
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconConsultar" ID="lnkConsultar" runat="server" Text="Consultar" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAjuda" ID="lnkAjudaConsulta" runat="server" Text="Ajuda" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Período:
                            </div>
                            <div class="coltxt" style="width: 140px;">
                                <asp:TextBox ID="txtDataInicial" runat="server" CssClass="calendario" Width="100px"
                                    data-ToolTip="default" ToolTip="Data inicial do período." />
                            </div>
                            <div class="coltxt" style="width: 140px;">
                                <asp:TextBox ID="txtDataFinal" runat="server" CssClass="calendario" Width="100px"
                                    data-ToolTip="default" ToolTip="Data final do período." />
                            </div>
                        </div>
                        <div class="bordagrid" style="width: 1000px">
                            <asp:GridView ID="gridF130Imobilizado" CssClass="compact" runat="server" AutoGenerateColumns="False"
                                CellPadding="4" radius-border="5px" ForeColor="#333333" GridLines="None"
                                OnSelectedIndexChanged="gridF130Imobilizado_SelectedIndexChanged">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <SelectedRowStyle BackColor="#e1e7ef" Font-Bold="True" ForeColor="#333333" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:CommandField SelectText="&gt;" ShowSelectButton="True" ButtonType="Button">
                                        <ItemStyle Width="25px" CssClass="tooltip"></ItemStyle>
                                    </asp:CommandField>
                                    <asp:BoundField DataField="MovimentoId" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Movimento">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Situacao" HeaderText="Situação">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="NatBcCred" HeaderText="NatBcCred">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="IdentBemImob" HeaderText="IdentBemImob">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="IndOrigCred" HeaderText="IndOrigCred">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="IndUtilBemImob" HeaderText="IndUtilBemImob">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="MesOpeAquis" HeaderText="MesOpeAquis">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="VlOperAquis" HeaderText="VlOperAquis">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ParcOperNaoBcCred" HeaderText="ParcOperNaoBcCred">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="VlBcCred" HeaderText="VlBcCred">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="IndNrParc" HeaderText="IndNrParc">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Cst" HeaderText="CstPis">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="AliqPis" HeaderText="AliqPis">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="VlPis" HeaderText="VlPis">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Cst" HeaderText="CstCofins">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="AliqCofins" HeaderText="AliqCofins">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="VlCofins" HeaderText="VlCofins">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="CodCta" HeaderText="CodCta">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="CodCcus" HeaderText="CodCcus">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="DescBemImob" HeaderText="DescBemImob">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaPlanoDeContas ID="ucConsultaPlanoDeContas" runat="server" />
</asp:Content>
