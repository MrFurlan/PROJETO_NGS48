<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="Requisicoes.aspx.vb" Inherits="NGS.Web.UI.Requisicoes" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngRequisicoes" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlRequisicoes" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Requisições
            </div>
            <ajaxToolkit:TabContainer ID="TabContainer1" runat="server" ActiveTabIndex="0" Width="100%">
                <ajaxToolkit:TabPanel runat="server" HeaderText="TabPanel1" ID="TabPanel1">
                    <HeaderTemplate>
                        Requisição
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li class="iconNovo" runat="server">
                                        <asp:LinkButton ID="lnkNovo" runat="server" Text="Gravar" />
                                    </li>
                                    <li class="iconExcluir" runat="server">
                                        <asp:LinkButton ID="lnkExcluir" runat="server" Text="Excluir" OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" />
                                    </li>
                                    <li class="iconConsultar" runat="server">
                                        <asp:LinkButton ID="lnkConsultar" runat="server" Text="Consultar" />
                                    </li>
                                    <li class="iconLimpar" runat="server">
                                        <asp:LinkButton ID="lnkLimpar" runat="server" Text="Limpar" />
                                    </li>
                                    <li class="iconAjuda" runat="server">
                                        <asp:LinkButton ID="lnkAjuda" runat="server" Text="Ajuda" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl ">
                                Unidade:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlUnidade" ToolTip="Unidade de negócio empresarial." runat="server" Width="595px" AutoPostBack="True" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl ">
                                Empresa:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlEmpresa" runat="server" Width="595px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl ">
                                Número:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtNumero" runat="server" Width="250px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl ">
                                Requisitante:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtRequisitante" runat="server" Width="495px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl ">
                                Depósito:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDeposito" runat="server" Width="495px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Solicitante:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtSolicitante" runat="server" Width="495px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Projeto:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlProjeto" runat="server" Width="595px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl ">
                                Setor:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlSetor" runat="server" Width="595px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl ">
                                C.Custo:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlCusto" runat="server" Width="595px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl ">
                                Data Entrega:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtdata" runat="server" Width="80px" CssClass="calendario" />
                            </div>
                            <div class="collbl">
                                Placa:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtPlaca" runat="server" Width="80px" CssClass="calendario" />
                            </div>
                            <div class="collbl">
                                Situação:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlSituação" runat="server" Width="300px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Observação:
                            </div>
                            <div class="coltxt" style="width: 85%;">
                                <asp:TextBox ID="txtObservacao" runat="server" Width="100%" TextMode="MultiLine" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="subtitulodiv">
                                Itens de Requisição:
                            </div>
                        </div>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li class="iconMais" runat="server">
                                        <asp:LinkButton ID="lnkAdicionar" runat="server" Text="Adicionar" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="grdRequisicao" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="grdRequisicao_SelectedIndexChanged">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:CommandField SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left" Width="25px"></ItemStyle>
                                    </asp:CommandField>
                                    <asp:BoundField DataField="Produto" HeaderText="Produto">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left" Width="60px"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Unidade" HeaderText="Unidade">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left" Width="60px"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Descricao" HeaderText="Descri&#231;&#227;o">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Observacao" HeaderText="Observa&#231;&#227;o">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left" Width="60px"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Quantidade" HeaderText="Quantidade">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left" Width="60px"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:CommandField HeaderText="Excluir" SelectText=" Excluir " ShowSelectButton="True"
                                        ButtonType="Button">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left" Width="50px"></ItemStyle>
                                    </asp:CommandField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" HeaderText="TabPanel2" ID="TabPanel2">
                    <HeaderTemplate>
                        Consulta
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="bordagrid">
                            <asp:GridView ID="grdConsulta" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="grdConsulta_SelectedIndexChanged">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:CommandField SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left" Width="25px"></ItemStyle>
                                    </asp:CommandField>
                                    <asp:BoundField DataField="Movimento" HeaderText="Movimento">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left" Width="60px"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Requisicao" HeaderText="Requisi&#231;&#227;o">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left" Width="60px"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Deposito" HeaderText="Deposito">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Requisitante" HeaderText="Requisitante">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left" Width="60px"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Setor" HeaderText="Setor">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left" Width="60px"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Solicitante" HeaderText="Solicitante">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left" Width="60px"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Situacao" HeaderText="Situa&#231;&#227;o">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left" Width="60px"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Projeto" HeaderText="Projeto">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left" Width="60px"></ItemStyle>
                                    </asp:BoundField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ProdutoNFG ID="ucProdutoNFG" runat="server" />
    <uc:ConsultaProduto ID="ucConsultaProduto" runat="server" />
</asp:Content>
