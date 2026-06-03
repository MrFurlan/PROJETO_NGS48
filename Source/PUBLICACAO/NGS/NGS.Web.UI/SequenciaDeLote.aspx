<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master" CodeBehind="SequenciaDeLote.aspx.vb" Inherits="NGS.Web.UI.SequenciaDeLote" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngSequenciaDeLote" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlSequenciaDeLote" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Produto X Sequência de Lote
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconNovo" ID="lnkNovo" Text="Gravar" runat="server" />
                        </li>
                        <li class="iconAtualizar" runat="server">
                            <asp:LinkButton ID="lnkAtualizar" Text="Atualizar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconExcluir" ID="lnkExcluir" Text="Excluir" runat="server"
                                OnClientClick="if(!confirm('Deseja realmente excluir o item selecionado?')) return false;" />
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
                <div class="collbl" style="width: 130px;">
                    Grupo:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlGrupoProduto" runat="server" Enabled="false" AutoPostBack="True" Width="576px" />
                </div>
                <div class="coltxt">
                    <asp:LinkButton ID="lnkBuscaProduto" runat="server" Enabled="false" Height="20px" Width="20px">
                        <asp:Image ID="imgSearch" ImageUrl="~/Images/search.png" runat="server" data-ToolTip="default"
                            ToolTip="Consultar Produto" />
                    </asp:LinkButton>
                </div>
            </div>
            <div class="row">
                <div class="collbl" style="width: 130px;">
                    Produto:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlProduto" runat="server" Enabled="false" AutoPostBack="True" Width="596px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl" style="width: 130px;">
                    Sequência do Produto:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtSequenciaProduto" runat="server" CssClass="txtNumerico3" Enabled="false" data-ToolTip="default" ToolTip="Sequência do Produto." />
                </div>
                <div class="collbl">
                    Sequência do Lote:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtSequenciaLote" runat="server" CssClass="txtNumerico4" Enabled="false" data-ToolTip="default" ToolTip="Sequência do Lote." />
                </div>
                <div class="collbl">
                    Ano:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtAno" runat="server" CssClass="txtNumerico4" Enabled="false" data-ToolTip="default" ToolTip="Ano do Lote." />
                </div>
            </div>
            <div class="row">
                <div class="collbl" style="width: 130px;">
                    Ordenar por:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdNome" Text="Nome" GroupName="cl" Checked="true"
                        runat="server" AutoPostBack="true" data-ToolTip="default" ToolTip="Ordena por Nome." />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdSequencia" Text="Sequência" GroupName="cl" 
                        runat="server" AutoPostBack="true" data-ToolTip="default" ToolTip="Ordena por sequência do Produto." />
                </div>
            </div>
            <div class="bordagrid" style="height: 525px;">
                <asp:GridView ID="gridSequenciaDeLote" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="gridSequenciaDeLote_SelectedIndexChanged">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True">
                            <HeaderStyle Width="50px" />
                            <ItemStyle Width="50px" />
                        </asp:CommandField>
                        <asp:BoundField DataField="Produto.Codigo" HeaderText="Produto" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" Width="100px" />
                            <ItemStyle HorizontalAlign="Left" Width="100px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Produto.Nome" HeaderText="Nome" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="SequenciaDoProduto" HeaderText="Sequência do Produto" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="SequenciaDoLote" HeaderText="Sequência do Lote" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Ano" HeaderText="Ano" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaProduto ID="ucConsultaProduto" runat="server" />
</asp:Content>
