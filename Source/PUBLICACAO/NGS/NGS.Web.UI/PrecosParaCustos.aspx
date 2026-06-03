<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="PrecosParaCustos.aspx.vb" Inherits="NGS.Web.UI.PrecosParaCustos" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngPrecosParaCustos" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlPrecosParaCustos" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Preços de Mercadorias para Custos
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
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:TextBox runat="server" ID="txtReduzidoEmp" data-ToolTip="default" title="Informe o código reduzido do depósito ou pressione a BARRA DE ESPAÇO para consultar."
                        Style="width: 50px;" TabIndex="5" MaxLength="50" />
                </div>
                <div class="coltxt">
                    <asp:TextBox runat="server" ID="txtCNPJEmp" data-ToolTip="default" title="Informe o CNPJ do depósito ou pressione a BARRA DE ESPAÇO para consultar."
                        Style="width: 115px;" TabIndex="6" MaxLength="50" />
                </div>
                <div class="coltxt">
                    <asp:TextBox runat="server" ID="txtCodigoEnderecoEmp" data-ToolTip="default" title="Informe o código do endereço do depósito."
                        Style="width: 30px;" TabIndex="7" MaxLength="50" />
                </div>
                <div class="coltxt">
                    <asp:TextBox runat="server" ID="txtNomeEmp" Style="width: 297px;" TabIndex="8" MaxLength="50"
                        size="23" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnPesquisarEmpresa" CssClass="btn" runat="server" UseSubmitBehavior="False"
                        Text=" > " data-ToolTip="default" ToolTip="Pesquisar Empresa" />
                </div>
                <div class="coltxt">
                    <img id="Img2" data-ToolTip="default" title="Limpar os campos" class="btn" alt="Limpar" src="images/Borracha.jpg" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Depósito:
                </div>
                <div class="coltxt">
                    <asp:TextBox runat="server" ID="txtReduzidoDep" data-ToolTip="default" title="Informe o código reduzido do depósito ou pressione a BARRA DE ESPAÇO para consultar."
                        Style="width: 50px;" TabIndex="5" MaxLength="50" />
                </div>
                <div class="coltxt">
                    <asp:TextBox runat="server" ID="txtCNPJDep" data-ToolTip="default" title="Informe o CNPJ do depósito ou pressione a BARRA DE ESPAÇO para consultar."
                        Style="width: 115px;" TabIndex="6" MaxLength="50" />
                </div>
                <div class="coltxt">
                    <asp:TextBox runat="server" ID="txtCodigoEnderecoDep" data-ToolTip="default" title="Informe o código do endereço do depósito."
                        Style="width: 30px;" TabIndex="7" MaxLength="50" />
                </div>
                <div class="coltxt">
                    <asp:TextBox runat="server" ID="txtNomeDep" Style="width: 297px;" TabIndex="8" MaxLength="50"
                        size="23" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnPesquisarDeposito" runat="server" class="btn" UseSubmitBehavior="False"
                        Text=" > " data-ToolTip="default" ToolTip="Pesquisar Depósito" />
                </div>
                <div class="coltxt">
                    <img id="Img4" data-ToolTip="default" title="Limpar os campos" class="btn" alt="Limpar" src="images/Borracha.jpg" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Grupo:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlGrupo" runat="server" Width="534px" AutoPostBack="true" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Produto:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlProduto" runat="server" Width="534px" />
                </div>
                 </div>
            <div class="row">
                <div class="collbl">
                    Data:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtData1" runat="server" CssClass="calendario" MaxLength="10" TabIndex="11"
                        Width="83px" />
                </div>
                <div class="collbl">
                    Preço:
                </div>
                <div class="coltxt">
                    <asp:TextBox runat="server" ID="txtPreco1" data-ToolTip="default" title="Informe o preço de mercado do produto."
                        Style="width: 100px;" MaxLength="30" CssClass=txtDecimal TabIndex="12" />
                </div>
                <div class="collbl">
                    Fator:
                </div>
                <div class="coltxt">
                    <asp:TextBox runat="server" ID="txtFator1" CssClass=txtDecimal data-ToolTip="default" title="Fator de conversão do peso." Style="width: 100px;"
                        TabIndex="13" />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="grdAbaTitulo" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True"
                            ItemStyle-HorizontalAlign="Center" ItemStyle-VerticalAlign="Middle" />
                        <asp:BoundField DataField="Reduzido" HeaderText="Reduzido" />
                        <asp:BoundField DataField="CNPJ" HeaderText="CNPJ" />
                        <asp:BoundField DataField="CodigoEndereco" HeaderText="End" />
                        <asp:BoundField DataField="Nome" HeaderText="Nome" />
                        <asp:BoundField DataField="Cidade" HeaderText="Cidade" />
                        <asp:BoundField DataField="UF" HeaderText="UF" />
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaEmpresas ID="ucConsultaEmpresas" runat="server" />
    <div id="divClientes" style="display: none; z-index: 101; left: 821px; width: 784px;
        position: absolute; top: 576px; background-color: #ffffff">
        <table class="consultaCli" id="table8" style="width: 100%; height: 216px; border: 0px none;">
            <tr class="clLinhaTabela2">
                <td style="width: 85px">
                    &nbsp;&nbsp;Reduzido
                </td>
                <td style="width: 128px">
                    &nbsp;CNPJ
                </td>
                <td style="width: 31px">
                    &nbsp;End
                </td>
                <td width="290">
                    &nbsp;Nome
                </td>
                <td style="width: 181px">
                    &nbsp;Cidade
                </td>
                <td style="width: 30px">
                    &nbsp;UF
                </td>
                <td style="height: 23px">
                    <img style="cursor: pointer" onclick="fecharDiv('divClientes');" height="21" alt=""
                        src="images/Fechar.jpg" width="21" />
                </td>
            </tr>
            <tr>
                <td style="width: 100%; height: 218px; vertical-align: text-top;" colspan="7">
                    <p style="text-align: left;">
                        <select class="fonteCourier" onkeypress="ClienteSelecionado();" style="width: 100%;
                            height: 100%" multiple size="12" name="ListaCli">
                        </select></p>
                </td>
            </tr>
        </table>
    </div>
    <div id="divTabelas" style="display: none; z-index: 102; left: 823px; width: 350px;
        position: absolute; top: 145px; height: 400px; background-color: #ffffff">
        <table id="table2" style="left: 0px; width: 336px; top: 248px; height: 406px; border: 1px none;">
            <tr class="tituloconsulta">
                <td style="width: 335px; height: 20px">
                    &nbsp;
                    <div id="divTitulo" style="display: inline; width: 304px; height: 15px; text-align: left;">
                        Label</div>
                </td>
                <td style="width: 23px;">
                    <img style="cursor: pointer" onclick="fecharDiv('divTabelas');" height="21" alt=""
                        src="images/Fechar.jpg" width="21">
                </td>
            </tr>
            <tr>
                <td colspan="2">
                    <p style="text-align: left;">
                        <select onkeypress="SelecionaTabela(event);" style="width: 340px; height: 376px"
                            multiple size="23" name="ListaOP">
                        </select></p>
                </td>
            </tr>
        </table>
    </div>
</asp:Content>
