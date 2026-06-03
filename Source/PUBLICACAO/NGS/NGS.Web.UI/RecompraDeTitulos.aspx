<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="RecompraDeTitulos.aspx.vb" Inherits="NGS.Web.UI.RecompraDeTitulos" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <script type="text/javascript">
        function selectAll(chkAll) {
            var chk = $('#' + chkAll.id);
            var checked = chk.attr('checked') == "checked";

            if (checked) {
                $("#MainContent_txtTotal").val(0);
            }

            $("input[type='checkbox']", "#MainContent_grdTitulos").not(".chkAlls").each(function () {
                $(this).attr("checked", checked);

                habilitaTxt(this);
            });
        };

        function habilitaTxt(ischk) {
            var chk = $('#' + ischk.id);
            var checked = chk.attr('checked') == "checked";

            var txt = $(ischk).parent().parent().find(".txt");
            var lbl = $("#MainContent_txtTotal");

            if (txt.val() != undefined) {

                txt.attr('disabled', !checked);

                if (lbl.val() != undefined) {

                    var total = parseFloat(lbl.val());

                    if (checked) {
                        total += parseFloat(txt.val().replace('.', '').replace(',', '.'));
                    }
                    else {
                        total -= parseFloat(txt.val().replace('.', '').replace(',', '.'));
                    }
                    lbl.val(total.toFixed(2));
                }
            }
        };

        function sumText() {
            var lbl = $("#MainContent_txtTotal");
            lbl.val(0);

            $("input[type='checkbox']", "#MainContent_grdTitulos").not(".chkAlls").each(function () {
                var checked = $(this).attr('checked') == "checked";
                var txt = $(this).parent().parent().find(".txt");

                var txtvalordoc = $(this).parent().parent().find(".txtValorDocumento");
                var txtvalorrec = $(this).parent().parent().find(".txtValorRecomprado");

                if (txt.val() != undefined) {
                    if (checked) {
                        var saldo = parseFloat(txtvalordoc.val().replace('.', '').replace(',', '.')) - parseFloat(txtvalorrec.val().replace('.', '').replace(',', '.'));

                        var total = parseFloat(lbl.val());
                        total += parseFloat(txt.val().replace('.', '').replace(',', '.'));

                        if (parseFloat(txt.val().replace('.', '').replace(',', '.')) > saldo) {
                            alert("O valor informado é maior do que o valor de saldo.");
                            //msgbox("O valor informado é maior do que o valor de saldo.", "ATENÇÃO!", "Info");
                        }
                        else {
                            lbl.val(total.toFixed(2));
                        }
                    }
                }
            });
        };
    </script>
    <style type="text/css">
        .txtTotal
        {
            text-align: right;
            margin-right: 4px;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmgRecompraDeTitulos" runat="server" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlRecompraDeTitulos" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Recompra de Títulos
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconNovo" runat="server" visible="false">
                            <asp:LinkButton ID="lnkNovo" Text="Gravar" runat="server" />
                        </li>
                        <li class="iconExcluir" runat="server" visible="false">
                            <asp:LinkButton ID="lnkExcluir" Text="Excluir" runat="server"
                                OnClientClick="if(!confirm('Deseja realmente excluir todos as recompras referente a este título?')) return false;" />
                        </li>
                        <li class="iconConsultar" runat="server">
                            <asp:LinkButton ID="lnkConsultar" Text="Consultar" runat="server" />
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li class="iconAjuda" runat="server">
                            <asp:LinkButton ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Título:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtTitulo" runat="server" CssClass="txtNumerico" Style="text-align: left;"
                        data-ToolTip="default" ToolTip="Informar o numero do titulo." />
                </div>
                <div class="collbl">
                    Contrato:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtContrato" runat="server" Enabled="False" CssClass="txtnumerico"
                        data-ToolTip="default" ToolTip="Informar o numero do contrato." />
                </div>
                <div class="collbl">
                    Valor:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtValorTP" runat="server" Enabled="False" CssClass="txtnumerico"
                        data-ToolTip="default" ToolTip="Valor do título." />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="grdTitulos" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    DataKeyNames="TituloRecomprar" ForeColor="#333333" GridLines="None" Width="100%">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:TemplateField>
                            <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="5px" />
                            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="5px" />
                            <HeaderTemplate>
                                <input id="chkAlls" onclick="selectAll(this);" type="checkbox" />
                            </HeaderTemplate>
                            <ItemTemplate>
                                <asp:CheckBox ID="chkSelecionado" onclick="habilitaTxt(this);" runat="server" />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:BoundField DataField="TituloRecomprar" HeaderText="Títulos (R)">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Empresa" HeaderText="Empresa">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Cliente" HeaderText="Cliente">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:TemplateField>
                            <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="5px" />
                            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="5px" />
                            <HeaderTemplate>
                                Valor Documento</HeaderTemplate>
                            <ItemTemplate>
                                <asp:TextBox ID="txtValorDocumento" Enabled="false" runat="server" Width="70px" CssClass="txtDecimal txtValorDocumento"
                                    Style="text-align: right;" Text='<%# bind("ValorDoDocumento", "{0:n2}") %>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField>
                            <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="5px" />
                            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="5px" />
                            <HeaderTemplate>
                                Valor Recomprado</HeaderTemplate>
                            <ItemTemplate>
                                <asp:TextBox ID="txtValorRecomprado" Enabled="false" runat="server" Width="70px"
                                    CssClass="txtDecimal txtValorRecomprado" Style="text-align: right;" Text='<%# bind("ValorRecomprado", "{0:n2}") %>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField>
                            <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="5px" />
                            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="5px" />
                            <ItemTemplate>
                                <asp:TextBox ID="txtValor" Enabled="false" onblur="sumText();" runat="server" Width="70px"
                                    CssClass="txtDecimal txt" Style="text-align: right;" Text='<%# bind("SaldoRecompra", "{0:n2}") %>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                    </Columns>
                </asp:GridView>
            </div>
            <div class="row">
                <div class="subtitulodiv">
                    <div class="coltxt" style="float: right;">
                        <asp:TextBox ID="txtTotal" runat="server" Enabled="false" CssClass="txtTotal" Height="20px"
                            Text="0.00" />
                    </div>
                    <div class="coltxt" style="float: right;">
                        <label>
                            SubTotal:
                        </label>
                    </div>
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
