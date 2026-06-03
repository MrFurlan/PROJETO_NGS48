<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="MovimentacoesFinanceiras.aspx.vb" Inherits="NGS.Web.UI.MovimentacoesFinanceiras" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngMovimentacoesFinanceiras" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="UpdatePanelPrincipal" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Movimentações Financeiras
            </div>
            <ajaxToolkit:TabContainer ID="TabContainer1" runat="server" ActiveTabIndex="0" Width="100%">
                <ajaxToolkit:TabPanel runat="server" HeaderText="TabPanel1" ID="TabPanel1">
                    <HeaderTemplate>
                        Titulos
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconNovo" ID="lnkNovo" runat="server" Text="Gravar" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconExcluir" ID="lnkExcluir" runat="server" Text="Excluir"
                                            OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconConsultar" ID="lnkConsultar" runat="server" Text="Consultar" />
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
                        <div class="row">
                            <div class="collbl">
                                Registro:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtRegistro" TabIndex="2" runat="server" Width="120px" Font-Size="10pt"
                                    Font-Bold="True" data-ToolTip="default" ToolTip="Número do lançamento." />
                            </div>
                            <div class="collbl">
                                Movimento:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtMovimento" TabIndex="1" CssClass="calendario" runat="server"
                                    Width="120px" Font-Size="9pt" Font-Bold="False" data-ToolTip="default" ToolTip="Data da consulta." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="subtitulodiv"> 
                                Empresa Solicitante
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Unidade:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="DdlUnidadeDeNegocioEmpresaCliente" TabIndex="3" runat="server"
                                    Width="623px" OnSelectedIndexChanged="DdlUnidadeDeNegocioEmpresaCliente_SelectedIndexChanged"
                                    AutoPostBack="True" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Empresa:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="DdlEmpresaCliente" TabIndex="4" runat="server" Width="623px"
                                    OnSelectedIndexChanged="DdlEmpresaCliente_SelectedIndexChanged" AutoPostBack="True" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Banco:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="DdlBancoFilial" TabIndex="5" runat="server" Width="623px" OnSelectedIndexChanged="DdlBancoFilial_SelectedIndexChanged"
                                    AutoPostBack="True" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Conta:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlContaFilial" runat="server" Width="623px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="subtitulodiv">
                                Empresa Pagadora
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Empresa Pagadora:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="DdlEmpresaPagadora" TabIndex="6" runat="server" Width="623px"
                                    OnSelectedIndexChanged="DdlEmpresaPagadora_SelectedIndexChanged1" AutoPostBack="True" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Tipo Pgto:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlTiposDePagamentos" runat="server" Width="623px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Banco:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="DdlBancoPagador" runat="server" Width="623px" OnSelectedIndexChanged="DdlBancoPagador_SelectedIndexChanged"
                                    AutoPostBack="True" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Conta:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="DdlContaPagadora" TabIndex="8" runat="server" Width="623px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Previsão Baixa:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="DdlProvisoes" TabIndex="9" runat="server" Width="359px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Histórico:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtHistorico" TabIndex="10" runat="server" Width="359px" TextMode="MultiLine"
                                    Height="20px" data-ToolTip="default" ToolTip="Selecionar o histórico adequado da operação." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Programação:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtVencimento" TabIndex="13" CssClass="calendario" runat="server"
                                    Width="120px" data-ToolTip="default" ToolTip="Data da programação do recebimento." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Valor:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtValor" CssClass="txtDecimal" TabIndex="19" runat="server" Width="120px"
                                    data-ToolTip="default" ToolTip="Inserir o valor total do pagamento." />
                            </div>
                            <div class="collbl">
                                Cheque:
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="rbChequeSim" runat="server" GroupName="cheque" Text="Sim" data-ToolTip="default"
                                    ToolTip="Informar o número da folha de cheque." />
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="rbChequeNao" runat="server" Checked="True" GroupName="cheque"
                                    Text="Não" data-ToolTip="default" ToolTip="Informar o número da folha de cheque." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Carteira:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlCarteiras" runat="server" Width="623px" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="BtnSlip" runat="server" Width="96px" Height="24px" Text="Slip" CssClass="btn"
                                    OnClick="BtnSlip_Click"/>
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" HeaderText="TabPanel4" ID="TabPanel4">
                    <HeaderTemplate>
                        Consulta Titulos
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconConsultar" ID="lnkConsultarCT" runat="server" Text="Consultar" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconLimpar" ID="lnkLimparCT" runat="server" Text="Limpar" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconRelatorio" ID="lnkRelatorioCT" runat="server" Text="Relatório" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAjuda" ID="lnkAjudaCT" runat="server" Text="Ajuda" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Unidade de Negócio:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="DdlUnidadeConsultaTitulos" runat="server" Width="596px" OnSelectedIndexChanged="DdlUnidadeConsultaTitulos_SelectedIndexChanged"
                                    AutoPostBack="True" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Empresa:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="DdlEmpresaConsultaTitulos" runat="server" Width="594px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Periodo Inicial:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtPeriodoInicialConsultaTitulos" class="calendario" runat="server"
                                    Width="116px" data-ToolTip="default" ToolTip="Periodo Inicial" />
                            </div>
                            <div class="collbl">
                                Periodo Final:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtPeriodoFinalConsultaTitulos" CssClass="calendario" runat="server"
                                    Width="108px" CausesValidation="True" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Opções Relatório:
                            </div>
                            <div class="coltxt">
                                <asp:Label ID="LbSituacao" runat="server" Font-Bold="True" Text="Situação:" />
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="RbGeral" runat="server" Text="Geral" OnCheckedChanged="RbGeral_CheckedChanged"
                                    Checked="True" GroupName="Situacao" />
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="RbAtivo" runat="server" Text="Ativo" OnCheckedChanged="RbAtivo_CheckedChanged"
                                    GroupName="Situacao" />
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="RbBaixado" runat="server" Text="Baixado" OnCheckedChanged="RbBaixado_CheckedChanged"
                                    GroupName="Situacao" />
                            </div>
                            <div class="coltxt">
                                <asp:Label ID="LbTotalizacao" runat="server" Font-Bold="True" Text="Totalizacao:" />
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="RbDiaGeral" runat="server" Width="67px" Text="Diario" OnCheckedChanged="RbDiaGeral_CheckedChanged"
                                    Checked="True" GroupName="Totalizacao" />
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="RbFilialDiario" runat="server" Text="Filial " OnCheckedChanged="RbFilialDiario_CheckedChanged"
                                    GroupName="Totalizacao" />
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="RbCarteiraDia" runat="server" Text="Carteira" OnCheckedChanged="RbCarteiraDia_CheckedChanged"
                                    GroupName="Totalizacao" />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="GridConsultaTitulos" runat="server" AutoGenerateColumns="False"
                                CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="GridConsultaTitulos_SelectedIndexChanged">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <EditRowStyle BackColor="#999999" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <Columns>
                                    <asp:CommandField SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button"
                                        HeaderText="Sel">
                                        <ItemStyle Width="30px"></ItemStyle>
                                    </asp:CommandField>
                                    <asp:BoundField DataField="Registro" HeaderText="Registro">
                                        <ItemStyle Width="60px"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Vencimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Vencimento"
                                        HtmlEncode="False">
                                        <ItemStyle Width="80px"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Cliente" HeaderText="Cliente">
                                        <ItemStyle Width="450px"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Valor" DataFormatString="{0:N}" HeaderText="Valor" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                </Columns>
                                <EmptyDataTemplate>
                                    <asp:CheckBox ID="CheckBox1" runat="server" />
                                </EmptyDataTemplate>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" HeaderText="TabPanel6" ID="TabPanel6">
                    <HeaderTemplate>
                        Observacoes
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="row">
                            <div class="coltxt">
                                <asp:TextBox ID="txtObservacoes" runat="server" Width="816px" Height="304px" TextMode="MultiLine" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="cmdAtualizarObservacoes" OnClick="cmdAtualizarObservacoes_Click"
                                    runat="server" Text="Atualizar" CssClass="btn" />
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
