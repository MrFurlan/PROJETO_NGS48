<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="CPR.aspx.vb" Inherits="NGS.Web.UI.CPR" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        .painelleft {
            width: 30%;
            margin-right: 6px;
        }

        .bordagrid {
            height: 200px;
        }
    </style>
    <script type="text/javascript" language="javascript">
        function pageLoadSelecaoProduto() {
            $("div.accordion").accordion({
                active: false,
                collapsible: true,
                alwaysOpen: false,
                heightStyle: "content",
                autoHeight: false,
                clearStyle: true
            });
        }

        $(document).ready(function () {
            pageLoadSelecaoProduto();
            var prmSelecaoProduto = Sys.WebForms.PageRequestManager.getInstance();
            prmSelecaoProduto.add_endRequest(pageLoadSelecaoProduto);
        });
    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngCPR" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlCPR" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                CPR - Cédula de Produto Rural
            </div>
            <ajaxToolkit:TabContainer ID="tcCPR" runat="server" ActiveTabIndex="0" Width="100%">
                <ajaxToolkit:TabPanel ID="tbCPR" runat="server" HeaderText="CPR">
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
                                        <asp:LinkButton class="iconRelatorio" ID="lnkRelatorio" Text="Relatório" runat="server" />
                                    </li>
                                    <li runat="server">
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
                                <asp:TextBox ID="txtEmpresa" runat="server" Enabled="False" Width="450px" />
                            </div>
                            <div class="coltxt">
                                <asp:Button runat="server" Text=">" CssClass="btn" UseSubmitBehavior="False" ID="btnEmpresa"
                                    OnClick="btnEmpresa_Click" data-ToolTip="default" ToolTip="Empresa Para negociação/consulta." />
                                <asp:HiddenField ID="txtCodigoEmpresa" runat="server" />
                            </div>
                            <div class="collbl">
                                Situação:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlSituacao" runat="server" Width="132px">
                                    <asp:ListItem Selected="True"></asp:ListItem>
                                    <asp:ListItem Value="1">Aberta</asp:ListItem>
                                    <asp:ListItem Value="0">Liquidada</asp:ListItem>
                                </asp:DropDownList>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Cartório:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtCartorio" runat="server" Enabled="False" Width="450px" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnCartorio" runat="server" Text=">" CssClass="btn" UseSubmitBehavior="False"
                                    OnClick="btnCartorio_Click" data-ToolTip="default" ToolTip="" />
                                <asp:HiddenField ID="txtCodigoCartorio" runat="server" />
                            </div>
                            <div class="collbl">
                                Data de Emissão:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDataEmissao" CssClass="calendario" runat="server" OnTextChanged="txtDataEmissao_TextChanged" data-ToolTip="default" ToolTip="Data de Criação da CPR." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Cliente:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtCliente" runat="server" Enabled="False" Width="450px" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnCliente" runat="server" Text=">" CssClass="btn" UseSubmitBehavior="False"
                                    OnClick="btnCliente_Click" data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                                <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                            </div>
                            <div class="collbl">
                                Registro:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtRegistro" runat="server" CssClass="txtdecimal" data-ToolTip="default" ToolTip="" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                CPR:
                            </div>
                            <div class="coltxt" style="width: 489px;">
                                <asp:TextBox ID="txtCPR" runat="server" data-ToolTip="default" ToolTip="" />
                            </div>
                            <div class="collbl">
                                Data do Registro:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDataRegistro" CssClass="calendario" runat="server"
                                    data-ToolTip="default" ToolTip="Data de registro da CPR" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Safra:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlSafra" runat="server" Width="489px" />
                            </div>
                            <div class="collbl">
                                Vencimento:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDataVencimento" CssClass="calendario" runat="server"
                                    data-ToolTip="default" ToolTip="Informar o vencimento da CPR" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Grupo:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlGrupoProduto" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlGrupoProduto_SelectedIndexChanged"
                                    Width="489px" />
                            </div>
                            <div class="collbl">
                                Quantidade:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtQtde" runat="server" CssClass="txtDecimal" AutoPostBack="True" data-ToolTip="default" ToolTip="Quantidade do produto." />
                            </div>
                            <div class="coltxt">
                                <asp:Label ID="lblProduto" Width="485px" runat="server" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Produto:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlProduto" runat="server" AutoPostBack="True" Width="459px"
                                    OnSelectedIndexChanged="ddlProduto_SelectedIndexChanged" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="BtnConsultaPrd" runat="server" OnClick="BtnConsultaPrd_Click" Text=" > "
                                    CssClass="btn" UseSubmitBehavior="False" data-ToolTip="default" ToolTip="Forma de compra/venda do produto" />
                            </div>
                            <div class="collbl">
                                Produtividade:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtProdutividade" runat="server" CssClass="txtDecimal" data-ToolTip="default" ToolTip="" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Endossado:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtEndosso" Width="450px" runat="server" Enabled="False" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnEndosso" runat="server" Text=">" CssClass="btn" UseSubmitBehavior="False" data-ToolTip="default" ToolTip="" />
                                <asp:HiddenField ID="txtCodigoEndosso" runat="server" />
                            </div>
                            <div class="collbl">
                                Área do CPR:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtAreaCPR" runat="server" CssClass="txtDecimal" AutoPostBack="True"
                                    Enabled="False" data-ToolTip="default" ToolTip="" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                CPR Cautela:
                            </div>
                            <div class="coltxt">
                                <asp:Label ID="lblDevedorSolidarioCPR" runat="server" />
                            </div>
                            <div class="coltxt">
                                <asp:Label ID="lblDescDevedorSolidario" runat="server" />
                            </div>
                            <div class="coltxt">
                                <asp:LinkButton ID="lnkDevedorSolidarioCPR" runat="server" CssClass="iconConsultar"
                                    OnClick="lnkDevedorSolidarioCPR_Click" data-ToolTip="default" ToolTip="" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Observação:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtObservacao" runat="server" TextMode="MultiLine" Width="743px" data-ToolTip="default" ToolTip="Inserir informação para observações relevantes." />
                            </div>
                        </div>
                        <div class="row" style="line-height: 14px;">
                            <ajaxToolkit:TabContainer ID="tcAuxiliar" runat="server" ActiveTabIndex="0" Width="100%">
                                <ajaxToolkit:TabPanel ID="tbAreaCPR" runat="server" HeaderText="Área">
                                    <ContentTemplate>
                                        <div class="painelleft">
                                            <div class="subtitulodiv">
                                                Fazenda
                                            </div>
                                            <div class="bordagrid">
                                                <asp:GridView ID="gridFazenda" runat="server" CellPadding="4" ForeColor="#333333"
                                                    GridLines="None" AutoGenerateColumns="False" OnSelectedIndexChanged="gridFazenda_SelectedIndexChanged"
                                                    OnRowDeleting="gridFazenda_RowDeleting" ShowFooter="True">
                                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                    <EditRowStyle BackColor="#999999" />
                                                    <Columns>
                                                        <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                                                        <asp:BoundField DataField="CodigoFazenda" HeaderText="Código">
                                                            <ItemStyle HorizontalAlign="Left" />
                                                            <HeaderStyle HorizontalAlign="Left" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="EndFazenda" HeaderText="End.">
                                                            <ItemStyle HorizontalAlign="Right" />
                                                            <HeaderStyle HorizontalAlign="Right" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="DescFazenda" HeaderText="Cliente/Fazenda">
                                                            <ItemStyle HorizontalAlign="Left" />
                                                            <HeaderStyle HorizontalAlign="Left" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="Area" HeaderText="Área" />
                                                        <asp:CommandField ButtonType="Image" DeleteImageUrl="~/images/ico-menos.gif" ShowDeleteButton="True" />
                                                    </Columns>
                                                </asp:GridView>
                                            </div>
                                        </div>
                                        <div class="painelleft">
                                            <div class="subtitulodiv">
                                                Matrículas
                                            </div>
                                            <div class="bordagrid">
                                                <asp:GridView ID="gridMatricula" runat="server" CellPadding="4" ForeColor="#333333"
                                                    GridLines="None" AutoGenerateColumns="False" OnRowDeleting="gridMatricula_RowDeleting">
                                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                    <EditRowStyle BackColor="#999999" />
                                                    <Columns>
                                                        <asp:BoundField DataField="CodigoMatricula" HeaderText="Matrícula" />
                                                        <asp:BoundField DataField="Area" HeaderText="Área" />
                                                        <asp:CommandField ButtonType="Image" DeleteImageUrl="~/images/ico-menos.gif" ShowDeleteButton="True" />
                                                    </Columns>
                                                </asp:GridView>
                                            </div>
                                        </div>
                                        <div class="painelleft">
                                            <div class="subtitulodiv">
                                                Proprietários
                                            </div>
                                            <div class="bordagrid">
                                                <asp:GridView ID="gridProprietario" runat="server" CellPadding="4" ForeColor="#333333"
                                                    GridLines="None" AutoGenerateColumns="False">
                                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                    <EditRowStyle BackColor="#999999" />
                                                    <Columns>
                                                        <asp:BoundField DataField="CodigoCliente" HeaderText="Proprietário" />
                                                        <asp:BoundField DataField="EndCliente" />
                                                        <asp:BoundField DataField="DescCliente" />
                                                    </Columns>
                                                </asp:GridView>
                                            </div>
                                        </div>
                                        <div class="painelleft" style="width: 5%; margin-top: 32px;">
                                            <asp:Button ID="btnAdicionarFazenda" runat="server" Text="Adicionar" CssClass="btn" />
                                        </div>
                                    </ContentTemplate>
                                </ajaxToolkit:TabPanel>
                                <ajaxToolkit:TabPanel ID="tbAvalista" runat="server" HeaderText="Avalista / Fiador">
                                    <ContentTemplate>
                                        <asp:Panel ID="pnlAvalista" runat="server" Height="400px" Width="100%">
                                            <div class="row">
                                                <div class="coltxt">
                                                    <asp:TextBox ID="txtAvalista" runat="server" Enabled="False" Width="500px" />
                                                </div>
                                                <div class="coltxt">
                                                    <asp:Button ID="btnAvalista" runat="server" Text=">" UseSubmitBehavior="False" CssClass="btn"
                                                        OnClick="btnAvalista_Click" />
                                                </div>
                                                <div class="coltxt">
                                                    <asp:DropDownList ID="ddlTipoRelacao" runat="server" Width="70px" />
                                                </div>
                                                <div class="coltxt">
                                                    <asp:Button ID="btnAdicionarAvalista" runat="server" Text="+" CssClass="btn" UseSubmitBehavior="False" />
                                                    <asp:HiddenField ID="txtCodigoAvalista" runat="server" />
                                                </div>
                                            </div>
                                            <div class="bordagrid" style="height: 365px">
                                                <asp:GridView ID="gridAvalista" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                                    ForeColor="#333333" GridLines="None" OnRowDeleting="gridAvalista_RowDeleting">
                                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                                    <EditRowStyle BackColor="#999999" />
                                                    <Columns>
                                                        <asp:BoundField DataField="CodigoAvalista" HeaderText="Código">
                                                            <ItemStyle HorizontalAlign="Left" />
                                                            <HeaderStyle HorizontalAlign="Left" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="EndAvalista" HeaderText="End.">
                                                            <HeaderStyle HorizontalAlign="Right" />
                                                            <ItemStyle HorizontalAlign="Right" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="DescAvalista" HeaderText="Avalista">
                                                            <ItemStyle HorizontalAlign="Left" />
                                                            <HeaderStyle HorizontalAlign="Left" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="DescTipoRelacao" HeaderText="Tipo de Relação">
                                                            <ItemStyle HorizontalAlign="Left" />
                                                            <HeaderStyle HorizontalAlign="Left" />
                                                        </asp:BoundField>
                                                        <asp:CommandField ButtonType="Image" DeleteImageUrl="~/images/ico-menos.gif" ShowDeleteButton="True" />
                                                    </Columns>
                                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                </asp:GridView>
                                            </div>
                                        </asp:Panel>
                                    </ContentTemplate>
                                </ajaxToolkit:TabPanel>
                                <ajaxToolkit:TabPanel ID="tbGrau" runat="server" HeaderText="Grau">
                                    <ContentTemplate>
                                        <asp:Panel ID="pnlGrau" runat="server" Height="400px" Width="100%">
                                            <div class="row">
                                                <div class="coltxt">
                                                    <asp:TextBox ID="txtGrau" runat="server" Enabled="False" Width="500px" />
                                                </div>
                                                <div class="coltxt">
                                                    <asp:Button ID="btnGrau" runat="server" Text=">" CssClass="btn" UseSubmitBehavior="False"
                                                        OnClick="btnGrau_Click" />
                                                </div>
                                                <div class="coltxt">
                                                    <asp:Button ID="btnAdicionarGrau" runat="server" Text="+" CssClass="btn" UseSubmitBehavior="False"
                                                        OnClick="btnAdicionarGrau_Click" />
                                                    <asp:HiddenField ID="txtCodigoGrau" runat="server" />
                                                </div>
                                            </div>
                                            <div class="bordagrid" style="height: 365px">
                                                <asp:GridView ID="gridGrau" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                                    ForeColor="#333333" GridLines="None" OnRowDeleting="gridGrau_RowDeleting">
                                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                                    <EditRowStyle BackColor="#999999" />
                                                    <Columns>
                                                        <asp:BoundField DataField="CodigoCliente" HeaderText="Código">
                                                            <HeaderStyle HorizontalAlign="Left" />
                                                            <ItemStyle HorizontalAlign="Left" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="EndCliente" HeaderText="End.">
                                                            <HeaderStyle HorizontalAlign="Right" />
                                                            <ItemStyle HorizontalAlign="Right" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="DescCliente" HeaderText="Cliente">
                                                            <ItemStyle HorizontalAlign="Left" />
                                                            <HeaderStyle HorizontalAlign="Left" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="Grau" HeaderText="Grau">
                                                            <ItemStyle HorizontalAlign="Right" />
                                                            <HeaderStyle HorizontalAlign="Right" />
                                                        </asp:BoundField>
                                                        <asp:CommandField ButtonType="Image" DeleteImageUrl="~/images/ico-menos.gif" ShowDeleteButton="True" />
                                                    </Columns>
                                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                </asp:GridView>
                                            </div>
                                        </asp:Panel>
                                    </ContentTemplate>
                                </ajaxToolkit:TabPanel>
                                <ajaxToolkit:TabPanel ID="tbLiquidacao" runat="server" HeaderText="Liquidação">
                                    <ContentTemplate>
                                        <asp:Panel ID="Panel1" runat="server" Height="400px" Width="100%">
                                            <div class="row">
                                                <div class="collbl">
                                                    Data:
                                                </div>
                                                <div class="coltxt">
                                                    <asp:TextBox ID="txtDataLiquidacao" runat="server" CssClass="calendario" Width="100px" />
                                                </div>
                                                <div class="collbl">
                                                    Quantidade:
                                                </div>
                                                <div class="coltxt">
                                                    <asp:TextBox ID="txtQtdLiquidacao" runat="server" CssClass="txtDecimal" Width="100px" />
                                                </div>
                                                <div class="coltxt">
                                                    <asp:Button ID="btnAdicionarLiquidacao" runat="server" Text="+" CssClass="btn" UseSubmitBehavior="False"
                                                        OnClick="btnAdicionarLiquidacao_Click" />
                                                </div>
                                            </div>
                                            <div class="bordagrid" style="height: 365px">
                                                <asp:GridView ID="GrdLiquidacao" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                                    ForeColor="#333333" GridLines="None" OnRowDeleting="grdLiquidacao_RowDeleting"
                                                    ShowFooter="True">
                                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                                    <EditRowStyle BackColor="#999999" />
                                                    <Columns>
                                                        <asp:BoundField DataField="Data" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Data">
                                                            <HeaderStyle HorizontalAlign="Left" />
                                                            <ItemStyle HorizontalAlign="Left" />
                                                        </asp:BoundField>
                                                        <asp:BoundField DataField="Quantidade" DataFormatString="{0:N2}" HeaderText="Quantidade">
                                                            <ItemStyle HorizontalAlign="Right" />
                                                            <HeaderStyle HorizontalAlign="Right" />
                                                        </asp:BoundField>
                                                        <asp:CommandField ButtonType="Image" DeleteImageUrl="~/images/ico-menos.gif" ShowDeleteButton="True" />
                                                    </Columns>
                                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                </asp:GridView>
                                            </div>
                                        </asp:Panel>
                                    </ContentTemplate>
                                </ajaxToolkit:TabPanel>
                            </ajaxToolkit:TabContainer>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="tbConsulta" runat="server">
                    <HeaderTemplate>
                        Consulta
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li class="iconConsultar" runat="server">
                                        <asp:LinkButton ID="lnkConsultarC" Text="Consultar" runat="server" />
                                    </li>
                                    <li class="iconLimpar" runat="server">
                                        <asp:LinkButton ID="lnkLimparC" Text="Limpar" runat="server" />
                                    </li>
                                    <li class="iconAjuda" runat="server">
                                        <asp:LinkButton ID="lnkAjudaC" Text="Ajuda" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div id="divConsultaCPR" class="accordion">
                            <h3>Consulta
                            </h3>
                            <div style="height: 100%; width: 100%;">
                                <div class="row">
                                    <div class="collbl">
                                        CPR:
                                    </div>
                                    <div class="coltxt" style="width: 236px;">
                                        <asp:TextBox ID="txtConsultaCPR" runat="server" Width="100px" data-ToolTip="default" ToolTip="" />
                                    </div>
                                    <div class="collbl">
                                        Registro:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtConsultaRegistro" runat="server" Width="100px" data-ToolTip="default" ToolTip="" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        Fazenda:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtFazenda" runat="server" Width="200px" Enabled="False" />
                                    </div>
                                    <div class="coltxt">
                                        <asp:Button ID="btnConsultaFazenda" runat="server" OnClick="btnConsultaFazenda_Click"
                                            Text=">" CssClass="btn" UseSubmitBehavior="False" data-ToolTip="default" ToolTip="Selecionar a fazenda desejada." />
                                        <asp:HiddenField ID="txtCodigoFazenda" runat="server" />
                                    </div>
                                    <div class="collbl">
                                        Matrículas:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtConsultaMatricula" runat="server" Width="100px" data-ToolTip="default" ToolTip="" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        Proprietário:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtProprietario" runat="server" Width="200px" Enabled="False" />
                                    </div>
                                    <div class="coltxt">
                                        <asp:Button ID="btnConsultaProprietario" runat="server" OnClick="btnConsultaProprietario_Click"
                                            Text=">" CssClass="btn" UseSubmitBehavior="False" data-ToolTip="default" ToolTip="Selecionar o proprietário da fazenda." />
                                        <asp:HiddenField ID="txtcodigoProprietario" runat="server" />
                                    </div>
                                    <div class="collbl">
                                        <asp:CheckBox ID="chkConsolidar" runat="server" Text="Consolidar:" data-ToolTip="default" ToolTip="Marcar para consolidar informações." />
                                    </div>
                                    <div class="coltxt" style="width: 100px">
                                        &nbsp;
                                    </div>
                                    <div class="collbl">
                                        Grau:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtConsultaGrau" runat="server" Width="26px" data-ToolTip="default" ToolTip="" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        Cliente:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtConsultaCliente" runat="server" Width="200px" Enabled="False" />
                                    </div>
                                    <div class="coltxt">
                                        <asp:Button ID="btnConsultaCliente" runat="server" Text=">" CssClass="btn" UseSubmitBehavior="False" data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                                        <asp:HiddenField ID="txtConsultaCodigoCliente" runat="server" />
                                    </div>
                                    <div class="collbl">
                                        Vencimento:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtDataVencimentoDe" CssClass="calendario" runat="server" Width="75px" data-ToolTip="default" ToolTip="Informar o período inicial e final de vencimento." />
                                    </div>
                                    <div class="collbl" style="width: 113px;">
                                        à:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtDataVencimentoAte" CssClass="calendario" runat="server" Width="75px" data-ToolTip="default" ToolTip="Informar o período inicial e final de vencimento." />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        Safra:
                                    </div>
                                    <div class="coltxt">
                                        <asp:DropDownList ID="ddlConsultaSafra" runat="server" Width="236px" />
                                    </div>
                                    <div class="collbl">
                                        Período de Emissão:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtDataEmissaoDe" CssClass="calendario" runat="server" Width="75px" data-ToolTip="default" ToolTip="Informar a data de inicio e fim da emissão." />
                                    </div>
                                    <div class="collbl" style="width: 113px;">
                                        à:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtDataEmissaoAte" CssClass="calendario" runat="server" Width="75px" data-ToolTip="default" ToolTip="Informar a data de inicio e fim da emissão." />
                                    </div>
                                </div>
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="gridConsulta" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Left" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                                    <asp:BoundField HeaderText="CodigoCartorio" DataField="CodigoCartorio">
                                        <HeaderStyle HorizontalAlign="Left" CssClass="Hide" />
                                        <ItemStyle HorizontalAlign="Left" CssClass="Hide" />
                                    </asp:BoundField>
                                    <asp:BoundField HeaderText="EndCartorio" DataField="EndCartorio">
                                        <HeaderStyle HorizontalAlign="Left" CssClass="Hide" />
                                        <ItemStyle HorizontalAlign="Left" CssClass="Hide" />
                                    </asp:BoundField>
                                    <asp:BoundField HeaderText="Cartorio" DataField="cartoriocidadeuf">
                                        <HeaderStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="CodigoCPR" HeaderText="CPR">
                                        <HeaderStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="DescricaoSituacao" HeaderText="Sit.">
                                        <HeaderStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="DataRegistro" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Dt. Registro">
                                        <HeaderStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="DataVencimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Vencimento">
                                        <HeaderStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="CodigoProduto" HeaderText="Cod.">
                                        <HeaderStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="NomeProduto" HeaderText="Nome Produto">
                                        <HeaderStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Quantidade" HeaderText="Qtde.">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Area" DataFormatString="{0:N2}" HeaderText="Área">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaEmpresas ID="ucConsultaEmpresas" runat="server" />
    <uc:ConsultaProduto ID="ucConsultaProduto" runat="server" />
    <uc:ConsultaFazendaCPR ID="ucConsultaFazendaCPR" runat="server" />
</asp:Content>
