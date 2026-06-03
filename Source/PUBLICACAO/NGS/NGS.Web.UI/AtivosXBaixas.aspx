<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="AtivosXBaixas.aspx.vb" Inherits="NGS.Web.UI.AtivosXBaixas" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngAtivosXbaixas" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlAtivosXbaixas" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Ativos X Baixas
            </div>
            <ajaxToolkit:TabContainer ID="TabContainer1" runat="server" Style="margin-top: 4px;"
                ActiveTabIndex="0" Width="100%">
                <ajaxToolkit:TabPanel runat="server" HeaderText="TabPanel1" ID="TabPanel1">
                    <HeaderTemplate>
                        Baixa
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconConsultar" ID="lnkConsultar" runat="server" Text="Consultar" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconNovo" ID="lnkBaixar" runat="server" Text="Baixar" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconLimpar" ID="lnkLimpar" runat="server" Text="Limpar" />
                                    </li>
                                    <li runat="server" style="width: 15%;">
                                        <asp:LinkButton class="iconExcluir" ID="lnkDesfazerBaixa" runat="server" OnClientClick="if(!confirm('Deseja realmente desfazer a baixa deste registro?')) return false;"
                                            Text="Desfazer Baixa" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconRelatorio" ID="lnkReemitirDocBaixa" runat="server" Text="Reemissão" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAjuda" ID="lnkAjuda" runat="server" Text="Ajuda" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Unidade:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlUnidade" runat="server" Width="590px" OnSelectedIndexChanged="ddlUnidade_SelectedIndexChanged"
                                    AutoPostBack="True" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Empresa:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="DdlEmpresaOrigem" runat="server" Width="590px" AutoPostBack="True" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Grupo:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="DdlGrupo" runat="server" Width="590px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Código:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtCodigo" runat="server" Width="107px" Enabled="False" MaxLength="4" data-ToolTip="default" ToolTip="Número do cadastro do bem." />
                            </div>
                            <div class="collbl">
                                Sequência:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtSequencia" runat="server" Width="86px" Enabled="False" MaxLength="4" data-ToolTip="default" ToolTip="De acordo com a quantidade de bens irá aumentar a sequência." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Aquisição:
                            </div>
                            <div class="coltxt" style="width: 117px;">
                                <asp:TextBox ID="txtAquisicao" CssClass="calendario" runat="server" Width="107px"
                                    Enabled="False" data-ToolTip="default" ToolTip="Data da compra do bem." />
                            </div>
                            <div class="collbl">
                                Inicio de uso:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtInicioDeUso" CssClass="calendario" runat="server" Width="86px"
                                    Enabled="False" data-ToolTip="default" ToolTip="Data de inicio da utilização do bem." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Descrição:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDescricao" runat="server" Width="580px" Enabled="False" data-ToolTip="default" ToolTip="Nome do produto." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Locação:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="DdlEmpresaDest" runat="server" Width="590px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Grupo de Conta:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="DdlGrupoDeContas" runat="server" Enabled="false" Width="590px"
                                    OnSelectedIndexChanged="DdlGrupoDeContas_SelectedIndexChanged" AutoPostBack="True" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Conta:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="DdlContaContabil" Enabled="false" runat="server" Width="590px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                C. Custo:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="DdlCentroDeCusto" Enabled="false" runat="server" Width="590px">
                                </asp:DropDownList>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Historico:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtHistorico" Enabled="false" runat="server" Width="580px" TextMode="MultiLine" data-ToolTip="default" ToolTip="Informações sobre o lançamento do bem." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Valor Original:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtValorOriginal" Enabled="false" class="txtDecimal" runat="server"
                                    Width="100px" data-ToolTip="default" ToolTip="Valor pago na aquisição do bem." />
                            </div>
                            <div class="collbl">
                                Depreciar:
                            </div>
                            <div class="coltxt" style="width: 90px;">
                                <asp:DropDownList ID="ddlDepreciar" Enabled="false" runat="server" AutoPostBack="True"
                                    Width="90px">
                                    <asp:ListItem Value="S">SIM</asp:ListItem>
                                    <asp:ListItem Value="N">NÃO</asp:ListItem>
                                </asp:DropDownList>
                            </div>
                            <div class="collbl">
                                <asp:CheckBox ID="chkSeguro" Enabled="false" runat="server" />Seguro:</span> <span
                                    style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Depreciado Inicial:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtValorDepreciadoInicial" Enabled="false" CssClass="txtDecimal"
                                    runat="server" Width="100px" />
                            </div>
                            <div class="collbl">
                                Atualização:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtAtualizacao" Enabled="false" CssClass="calendario" runat="server"
                                    Width="80px" data-ToolTip="default" ToolTip="Data da ultima atualização do bem." />
                            </div>
                            <div class="collbl">
                                Data Baixa:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDataBaixa" CssClass="calendario" runat="server" Width="78px" data-ToolTip="default" ToolTip="Data que ocorrer a baixa do bem." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Depreciado Total:
                            </div>
                            <div class="coltxt" style="margin-right: 224px;">
                                <asp:TextBox ID="txtValorDepreciado" runat="server" Enabled="False" Width="100px" data-ToolTip="default" ToolTip="Valor já depreciado." />
                            </div>
                            <div class="collbl">
                                Valor da Baixa:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtValorBaixa" CssClass="txtDecimal" runat="server" Width="78px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Motivo da Baixa.:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtMotivoBaixa" runat="server" Width="579px" TextMode="MultiLine" data-ToolTip="default" ToolTip="Informar qual o motivo para baixa do bem." />
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" HeaderText="TabPanel2" ID="TabPanel2">
                    <HeaderTemplate>
                        Consulta
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="bordagrid">
                            <asp:GridView ID="GridAtivos" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="GridAtivos_SelectedIndexChanged">
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
                                    <asp:BoundField DataField="Grupo" HeaderText="Grupo">
                                        <ItemStyle Width="50px"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Codigo" HeaderText="C&#243;digo">
                                        <ItemStyle Width="50px"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Sequencia" HeaderText="Sequ&#234;ncia">
                                        <ItemStyle Width="50px"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Descricao" HeaderText="Descricao">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Aquisicao" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Aquisi&#231;&#227;o">
                                        <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Center" Width="60px"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Valor" DataFormatString="{0:N}" HeaderText="Valor">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right" Width="100px"></ItemStyle>
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
