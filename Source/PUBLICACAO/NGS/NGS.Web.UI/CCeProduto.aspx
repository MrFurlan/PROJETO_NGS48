<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="CCeProduto.aspx.vb" Inherits="NGS.Web.UI.CCeProduto" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        #meioconteudo {
            width: 1180px !important;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngCCeObservacao" runat="server" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlCCeObservacao" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                CCe Produto (NFe)
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconAtualizar" ID="lnkAtualizar" runat="server" Text="Atualizar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconConsultar" ID="lnkConsultar" runat="server" Text="Consultar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconRelatorio" ID="lnkReimpressao" Text="Reimpressão" runat="server" />
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
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUnidadeDeNegocio" runat="server" Width="630px" AutoPostBack="True"
                        OnSelectedIndexChanged="ddlUnidadeDeNegocio_SelectedIndexChanged" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="630px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                    <asp:TextBox ID="txtCliente" runat="server" Width="591px" Enabled="False" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnCliente" runat="server" Text=" > " UseSubmitBehavior="False" CssClass="btn"
                        OnClick="btnCliente_Click" data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data Inicial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicial" runat="server" CssClass="calendario" Width="100px"
                        data-ToolTip="default" ToolTip="Período inicial da pesquisa." />
                </div>
                <div class="collbl">
                    Data Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" runat="server" CssClass="calendario" Width="100px"
                        data-ToolTip="default" ToolTip="Período final da pesquisa." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    E/S:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtES" runat="server" Width="100px" MaxLength="1" data-ToolTip="default"
                        ToolTip="" />
                </div>
                <div class="collbl" style="margin-left: 21px;">
                    Nota:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNota" Style="text-align: right;" runat="server" class="txtNumerico"
                        Width="100px" data-ToolTip="default" ToolTip="Número da Nota Fiscal." />
                </div>
                <div class="collbl">
                    Série:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtSerie" runat="server" Width="100px" MaxLength="3" data-ToolTip="default"
                        ToolTip="Série da Nota Fiscal." />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="grdCceProdutos" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:BoundField DataField="Sequencia" HeaderText="Seq">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left" Width="30px"></ItemStyle>
                        </asp:BoundField>
                        <asp:TemplateField>
                            <HeaderTemplate>
                                Produto
                            </HeaderTemplate>
                            <ItemTemplate>
                                <asp:CheckBox ID="chkProduto" runat="server" AutoPostBack="True" OnCheckedChanged="chkProduto_CheckedChanged" />
                                <asp:TextBox ID="txtProduto" Enabled="false" Text='<%# Bind("CodigoProduto")%>' runat="server"
                                    Width="80px" />
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:TemplateField>
                        <asp:TemplateField>
                            <HeaderTemplate>
                                Descrição
                            </HeaderTemplate>
                            <ItemTemplate>
                                <asp:TextBox ID="txtNomeProduto" Enabled="false" Text='<%# Bind("Produto.Nome")%>' runat="server" TextMode="MultiLine" Width="250px" />
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:TemplateField>
                        <asp:TemplateField>
                            <HeaderTemplate>
                                Unidade
                            </HeaderTemplate>
                            <ItemTemplate>
                                <asp:TextBox ID="txtUnidade" Enabled="false" Width="50px" Text='<%# Bind("Produto.Unidade")%>' runat="server" />
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:TemplateField>
                        <asp:TemplateField>
                            <HeaderTemplate>
                                NCM
                            </HeaderTemplate>
                            <ItemTemplate>
                                <asp:CheckBox ID="chkNCM" runat="server" />
                                <asp:TextBox ID="txtNCM" Enabled="false" Text='<%# Bind("produto.NCM")%>' runat="server"
                                    Width="60px" />
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" Width="90px" />
                        </asp:TemplateField>
                        <asp:BoundField DataField="QuantidadeFiscal" DataFormatString="{0:N0}" HeaderText="Quantidade">
                            <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Unitario" DataFormatString="{0:N2}" HeaderText="Unitário">
                            <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="ValorTotal" DataFormatString="{0:N2}" HeaderText="Valor Total">
                            <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                        </asp:BoundField>
                        <asp:TemplateField>
                            <HeaderTemplate>
                                Observação
                            </HeaderTemplate>
                            <ItemTemplate>
                                <table style="width: 100%; border: 0px none;" cellpadding="0" cellspacing="0">
                                    <tr>
                                        <td>
                                            <asp:CheckBox ID="chkObservacaoProd" runat="server" OnCheckedChanged="chkObservacaoProd_CheckedChanged"
                                                AutoPostBack="true" />
                                        </td>
                                        <td>
                                            <asp:TextBox ID="txtObservacaoProd" Enabled="false" Text='<%# bind("ObservacoesDoProduto") %>'
                                                runat="server" TextMode="MultiLine" Width="250px" />
                                        </td>
                                    </tr>
                                </table>
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:TemplateField>
                        <asp:TemplateField>
                            <HeaderTemplate>
                                OpXEnc
                            </HeaderTemplate>
                            <ItemTemplate>
                                <asp:TextBox ID="txtOperacaoEstado" Enabled="false" Text='<%# Bind("CodigoOperacaoEstado")%>' runat="server"  Width="50px" />
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:TemplateField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaPedidosxNotas ID="ucConsultaPedidosXNotas" runat="server" />
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaProduto ID="ucConsultaProduto" runat="server" />
</asp:Content>
