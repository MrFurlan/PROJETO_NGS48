<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucMonitorCupomFiscal.ascx.vb"
    Inherits="NGS.Web.UI.ucMonitorCupomFiscal" %>
<script type="text/javascript">
    function pageLoadMonitorCupomFiscal() {
        $("#MainContent_ucMonitorCupomFiscal_btnFechar", "#divMonitorCupomFiscal").button();
    }

    $(document).ready(function () {
        pageLoadMonitorCupomFiscal();
    });

    var prmMonitorCupomFiscal = Sys.WebForms.PageRequestManager.getInstance();
    prmMonitorCupomFiscal.add_endRequest(pageLoadMonitorCupomFiscal);
</script>
<div id="divMonitorCupomFiscal" class="uc" title="Monitor de Cupom Fiscal" style="display: none;">
    <asp:UpdatePanel ID="updpnlMonitorDeNotas" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <asp:HiddenField ID="hdfMovimento" runat="server" />
            <script type="text/javascript">
                function selectAllCheckboxes(chkAll) {
                    var chk = $('#' + chkAll.id);
                    var checked = chk.attr('checked') == "checked";
                    $("input[type='checkbox']", "#MainContent_ucMonitorCupomFiscal_GridNotas").not("#chkAll").each(function () {
                        if (!$(this).is(':disabled')) {
                            $(this).attr("checked", checked);
                        }
                    });
                }
            </script>
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td>
                        <table class="actions" style="width: 100%;">
                            <tr>
                                <td class="iconMail" runat="server" style="width: 10%;">
                                    <asp:LinkButton ID="lnkEnviarEmail" runat="server">
                                        <span>Enviar E-mail</span>
                                    </asp:LinkButton>
                                </td>
                                <td class="iconRelatorio" runat="server" style="width: 10%;">
                                    <asp:LinkButton ID="lnkImprimir" runat="server">
                                        <span>Imprimir</span>
                                    </asp:LinkButton>
                                </td>
                                <td runat="server" style="width: 60%;" align="right">
                                    <label>
                                        Situação:
                                    </label>
                                    <asp:DropDownList ID="ddlSituacao" runat="server" Width="200px" AutoPostBack="true"
                                        OnSelectedIndexChanged="ddlSituacao_SelectedIndexChanged" />
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr>
                    <td>
                        <div class="bordagrid" style="min-height: 400px; max-height: 400px; max-width: 100%; width: 100%; overflow: auto;">
                            <asp:GridView ID="GridNotas" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="GridNotas_SelectedIndexChanged">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                                    <asp:TemplateField>
                                        <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="5px" />
                                        <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="5px" />
                                        <HeaderTemplate>
                                            <input id="chkAll" onclick="selectAllCheckboxes(this);" type="checkbox" />
                                        </HeaderTemplate>
                                        <ItemTemplate>
                                            <asp:CheckBox ID="chkSelecionado" runat="server" Enabled="False" />
                                            <asp:HiddenField ID="hdfKey" runat="server" Value='<%# String.Format("{0}-{1}-{2}-{3}-{4}-{5}-{6}", Eval("Empresa_Id"), Eval("EndEmpresa_Id"), Eval("Cliente"), Eval("EndCliente"), Eval("ES"), Eval("Nota"), Eval("Serie")) %>' />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:Image ID="imgStatus" runat="server" ImageUrl='<%# GetImagem(Eval("Retorno")) %>'
                                                ToolTip='<%# GetTooltip(String.Format("{0} - {1}", Eval("Retorno"), Eval("MsgRetorno"))) %>' />
                                        </ItemTemplate>
                                        <HeaderStyle HorizontalAlign="Center" />
                                        <ItemStyle HorizontalAlign="Center" />
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="Movimento" HeaderText="Movimento" DataFormatString="{0:dd/MM/yyyy}">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Empresa_id" HeaderText="Empresa">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="EndEmpresa_id" HeaderText="End">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Cliente" HeaderText="Cliente" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="EndCliente" HeaderText="End">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="NomeCliente" HeaderText="Nome">
                                        <HeaderStyle HorizontalAlign="Left" Width="350px" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ES" HeaderText="E/S">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Serie" HeaderText="S&#233;rie">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Nota" HeaderText="Nota">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Operacao" HeaderText="OP">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="SubOperacao" HeaderText="SO">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Quantidade" HeaderText="Quantidade" DataFormatString="{0:n0}">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" Width="100px" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Unitario" HeaderText="Unit&#225;rio" DataFormatString="{0:n6}">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" Width="100px" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Valor" HeaderText="Valor" DataFormatString="{0:n2}">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" Width="100px" />
                                    </asp:BoundField>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:HiddenField ID="hdfNossaEmissao" runat="server" Value='<%# Eval("NossaEmissao") %>' />
                                            <asp:LinkButton ID="lnkConsultar" runat="server" OnClick="lnkConsultar_Click">
                                                <img src='<%# ResolveUrl("~/Images/find.png") %>' alt="Consultar NFC-e" data-ToolTip="default" title="Consultar NFC-e" /> 
                                            </asp:LinkButton>
                                        </ItemTemplate>
                                        <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                        <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </td>
                </tr>
                <tr>
                    <td align="right" style="padding: 10px 0px 0px 10px;">
                        <asp:Button ID="btnFechar" runat="server" Text="Fechar" />
                    </td>
                </tr>
            </table>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
