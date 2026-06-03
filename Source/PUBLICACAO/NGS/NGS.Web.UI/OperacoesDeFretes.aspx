<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="OperacoesDeFretes.aspx.vb" Inherits="NGS.Web.UI.OperacoesDeFretes" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngOperacoesDeFretes" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlOperacoesDeFretes" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Operações de Fretes
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
                        <li class="iconExcluir" runat="server">
                            <asp:LinkButton ID="lnkExcluir" Text="Excluir" runat="server" OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" />
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
                    Operação:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlOperacao" runat="server" Width="690px" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Sub-Operação:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlSubOperacao" runat="server" Width="690px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Op. Destino:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlOpDestino" runat="server" Width="690px" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Sub-Op. Destino:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlSubOpDestino" runat="server" Width="690px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Op. Anulação:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlOpAnulacao" runat="server" Width="690px" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Sub-Op. Anulação:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlSubOpAnulacao" runat="server" Width="690px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Op. Contrapartida:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlOpContrapartida" runat="server" Width="690px" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Sub-Op. Contrap:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlSubOpContrapartida" runat="server" Width="690px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Tipo de Pessoa:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdoFisica" runat="server" Text=" Física " GroupName="grpTipoPessoa"
                        Checked="true" data-ToolTip="default" ToolTip="Selecionar se é pessoa física ou juídica." />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdoJuridica" runat="server" Text=" Jurídica " GroupName="grpTipoPessoa"
                        data-ToolTip="default" ToolTip="Selecionar se é pessoa física ou juídica." />
                </div>
                <div class="collbl">
                    Tipo de Operação:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdoCirculacao" runat="server" Text=" Circulação " GroupName="grpTipoOperacao"
                        Checked="true" AutoPostBack="True" data-ToolTip="default" ToolTip="Selecionar se é refernete a circulação, prestação de serviço ou estadia." />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdoPrestacaoServico" runat="server" Text=" Prestação de Serviço "
                        GroupName="grpTipoOperacao" AutoPostBack="True" data-ToolTip="default" ToolTip="Selecionar se é refernete a circulação, prestação de serviço ou estadia." />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdoEstadia" runat="server" Text=" Estadia " GroupName="grpTipoOperacao"
                        AutoPostBack="True" data-ToolTip="default" ToolTip="Selecionar se é refernete a circulação, prestação de serviço ou estadia." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Formação De Lote:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlFormacaoDeLote" runat="server" Width="110px">
                        <asp:ListItem Value="0">Não</asp:ListItem>
                        <asp:ListItem Value="1">Sim</asp:ListItem>
                    </asp:DropDownList>
                </div>
                <div runat="server" id="divTipoServico" visible="false">
                    <div class="collbl">
                        Prestação de Serviço:
                    </div>
                    <div class="coltxt">
                        <asp:RadioButton ID="rdoPropria" runat="server" Text="Própria" GroupName="grpTipoServico"
                            Checked="true" />
                        <asp:RadioButton ID="rdoTerceiros" runat="server" Text="Terceiros" GroupName="grpTipoServico" />
                    </div>
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="grd" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="grd_SelectedIndexChanged">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:CommandField SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button">
                            <ItemStyle Width="25px"></ItemStyle>
                        </asp:CommandField>
                        <asp:BoundField DataField="OperacaoId" HeaderText="Op." HeaderStyle-HorizontalAlign="Center"
                            ItemStyle-HorizontalAlign="Center" />
                        <asp:BoundField DataField="SubOperacaoId" HeaderText="Sub-Op." HeaderStyle-HorizontalAlign="Center"
                            ItemStyle-HorizontalAlign="Center" />
                        <asp:BoundField DataField="EntradaSaida" HeaderText="E/S" HeaderStyle-HorizontalAlign="Center"
                            ItemStyle-HorizontalAlign="Center" />
                        <asp:BoundField DataField="TipoPessoaId" HeaderText="Tipo de Pessoa" HeaderStyle-HorizontalAlign="Center"
                            ItemStyle-HorizontalAlign="Center" />
                        <asp:BoundField DataField="Classe" HeaderText="Classe" HeaderStyle-HorizontalAlign="Center"
                            ItemStyle-HorizontalAlign="Center" />
                        <asp:BoundField DataField="OpDestinoId" HeaderText="Op. Destino" HeaderStyle-HorizontalAlign="Center"
                            ItemStyle-HorizontalAlign="Center" />
                        <asp:BoundField DataField="SubOpDestinoId" HeaderText="Sub Destino" HeaderStyle-HorizontalAlign="Center"
                            ItemStyle-HorizontalAlign="Center" />
                        <asp:BoundField DataField="OpAnulacao" HeaderText="Op. Anulação" HeaderStyle-HorizontalAlign="Center"
                            ItemStyle-HorizontalAlign="Center" />
                        <asp:BoundField DataField="SubOpAnulacao" HeaderText="Sub Anulação" HeaderStyle-HorizontalAlign="Center"
                            ItemStyle-HorizontalAlign="Center" />
                        <asp:BoundField DataField="OpContrapartida" HeaderText="Op. Contrapartida" HeaderStyle-HorizontalAlign="Center"
                            ItemStyle-HorizontalAlign="Center" />
                        <asp:BoundField DataField="SubOpContrapartida" HeaderText="Sub Contrapartida" HeaderStyle-HorizontalAlign="Center"
                            ItemStyle-HorizontalAlign="Center" />
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
