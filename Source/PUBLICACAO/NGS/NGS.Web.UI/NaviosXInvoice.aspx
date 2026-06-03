<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master" CodeBehind="NaviosXInvoice.aspx.vb" Inherits="NGS.Web.UI.NaviosXInvoice" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <script type="text/javascript">

        function pageLoadChegadaNavios() {

            $("#txtDataDeChegada:enabled").datepicker({
                dateFormat: 'dd/mm/yy',
                dayNames: ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado', 'Domingo'],
                dayNamesMin: ['D', 'S', 'T', 'Q', 'Q', 'S', 'S', 'D'],
                dayNamesShort: ['Dom', 'Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb', 'Dom'],
                monthNames: ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'],
                monthNamesShort: ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'],
                nextText: 'Próximo',
                prevText: 'Anterior',
                showOn: "button",
                buttonImage: "Images/calendar.png",
                buttonImageOnly: true,
                showButtonPanel: true
            });

            $("#txtDataDeChegada").setMask("date");
        }

        $(document).ready(function () {
            pageLoadChegadaNavios();
            var prmChegadaNavios = Sys.WebForms.PageRequestManager.getInstance();
            prmChegadaNavios.add_endRequest(pageLoadChegadaNavios);
        });
    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngBancos" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlBancos" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Invoice
            </div>
            <ajaxToolkit:TabContainer ID="TabNavio" runat="server" ActiveTabIndex="0" Width="100%" Style="clear: both;">
                <ajaxToolkit:TabPanel runat="server" ID="TabNavioXProduto">
                    <HeaderTemplate>
                        Produto
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li class="iconNovo" runat="server">
                                        <asp:LinkButton ID="lnkNovo" runat="server" Text="Gravar" />
                                    </li>
                                    <li class="iconAtualizar" runat="server">
                                        <asp:LinkButton ID="lnkAtualizar" runat="server" Text="Atualizar" />
                                    </li>
                                    <li class="iconExcluir" runat="server">
                                        <asp:LinkButton ID="lnkExcluir" runat="server" Text="Desativar"
                                            OnClientClick="if(!confirm('Deseja realmente desativar este registro?')) return false;" />
                                    </li>
                                    <li class="iconLimpar" runat="server">
                                        <asp:LinkButton ID="lnkLimpar" runat="server" Text="Limpar" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAjuda" ID="lnkAjuda" Text="Ajuda" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Invoice:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtCodigo" runat="server" Width="76px" data-ToolTip="default" ToolTip="Código invoice." />
                            </div>
                            <div class="collbl">
                                Navio:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtNavio" runat="server" Width="70px" data-ToolTip="default" ToolTip="Código do navio." />
                                <asp:LinkButton class="iconConsultar" ID="lnkConsultarNavio" runat="server" Enabled="False" />
                            </div>
                            <div class="collbl">
                                Descrição:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDescricao" runat="server" Width="300px" MaxLength="390" data-ToolTip="default"
                                    ToolTip="Descrição do navio." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbluc">
                                Grupo:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlGrupoProduto" runat="server" Enabled="False" AutoPostBack="True" Width="576px" />
                            </div>
                            <div class="coltxt">
                                <asp:LinkButton ID="lnkBuscaProduto" runat="server" Enabled="False" Height="20px" Width="20px"><asp:Image ID="imgSearch" ImageUrl="~/Images/search.png" runat="server" data-ToolTip="default"
                                        ToolTip="Consulta Produto" /></asp:LinkButton>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbluc">
                                Produto:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlProdutos" runat="server" Enabled="False" AutoPostBack="True" Width="596px" />
                            </div>
                        </div>
                        <div>
                            <div class="collbluc">
                                Origem do Produto:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlPais" runat="server" Enabled="False" AutoPostBack="True" Width="596px"  />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Observação:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtObservacao" runat="server" TextMode="MultiLine" Width="587px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Data chegada:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDataDeChegada" CssClass="calendario" runat="server" Enabled="True" data-ToolTip="default"  ToolTip="Data de chegada." />
                            </div>
                            <div>
                                <div class="collbl">
                                    Ativo:
                                </div>
                                <asp:CheckBox ID="chkAtivo" runat="server" Text="Ativo" />
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" ID="TabNavios">
                    <HeaderTemplate>
                        Navios
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton ID="lnkConsultar" CssClass="iconConsultar" Text="Consultar"
                                            runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconRelatorio" ID="lnkRelatorio" Text="Relatório" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="GridNaviosXInvoice" runat="server" AutoGenerateColumns="False"
                                CellPadding="4" ForeColor="#333333" Width="100%" GridLines="None">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:CommandField SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left" Width="40px"></ItemStyle>
                                    </asp:CommandField>
                                    <asp:BoundField DataField="Codigo" HeaderText="Invoice">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left" Width="50px"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Descricao" HeaderText="Navio">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Observacao" HeaderText="Observa&#231;&#227;o">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="DataDeChegada" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Data de chegada">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Pais.Descricao" HeaderText="Origem">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="NavioxinvoiceXProduto.DescricaoProduto" HeaderText="Nome">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:TemplateField HeaderText="Ativo">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemTemplate>
                                            <asp:CheckBox ID="chkAtivogrid" runat="server" Checked='<%# Bind("Ativo") %>' Enabled="false" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultarNavios ID="ucConsultarNavios" runat="server" />
    <uc:ConsultaProduto ID="ucConsultaProduto" runat="server" />
</asp:Content>
