<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="NotaFiscalXClassificacao.aspx.vb"
    Inherits="NGS.Web.UI.NotaFiscalXClassificacao" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Classificação Nota Fiscal</title>
    <script type="text/javascript">
        function iniciar() { }

        function fecharJanela() {
            var theForm = window.opener.document.forms[0];
            window.close();
        }

        function ForcarPostBack() {
            var theForm = window.opener.document.forms[0];
            theForm.submit();
        }

        function ValidarObservacao(event) {
            if (document.all) {
                event.returnValue = false;
            }
        }

        function MudarFoco(tecla, tamanho, obrigatorio, anterior, proximo) {
            var validacao = (obrigatorio ? tamanho > 0 : true);
            switch (tecla) {
                case 13:
                    proximo.focus();
                    break;
                case 27:
                    anterior.focus();
                    break;
                default:
                    return false;
            }

            return true;
        }	   
    </script>
</head>
<body onunload="ForcarPostBack();">
    <form id="form1" runat="server">
    <div>
        <asp:ScriptManager ID="ScriptManager1" runat="server" EnableScriptGlobalization="True"
            EnableScriptLocalization="True"/>
        <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
        <asp:UpdatePanel ID="UpdatePanel1" runat="server">
            <ContentTemplate>
                <table class="borda" width="100%">
                    <tr>
                        <td class="titulotabela" colspan="3" bgcolor="#e8e8e8" align="center">
                            Classificação Nota Fiscal<asp:HiddenField ID="HID" runat="server" />
                        </td>
                    </tr>
                    <tr>
                        <td class="titulotabela" colspan="2" align="center">
                            <asp:Label ID="lblEntSai" runat="server" ForeColor="Red"/>
                            <asp:HiddenField ID="txtEntradaSaida" runat="server" />
                            <asp:HiddenField ID="txtProduto" runat="server" />
                        </td>
                        <td rowspan="6" align="left" valign="top">
                            <asp:GridView ID="GridDescontos" runat="server" AutoGenerateColumns="False" BackColor="White"
                                BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px" CellPadding="1" SelectedIndex="170"
                                Width="100%">
                                <FooterStyle BackColor="White" ForeColor="#000066" />
                                <RowStyle ForeColor="#000066" />
                                <Columns>
                                    <asp:BoundField DataField="Codigo" HeaderText="Codigo" />
                                    <asp:BoundField DataField="Descricao" HeaderText="Descri&#231;&#227;o" />
                                    <asp:TemplateField HeaderText="Percentual">
                                        <ItemTemplate>
                                            <ajaxToolkit:MaskedEditExtender ID="mskPercentual" runat="server" Mask="99.99" MaskType="Number"
                                                TargetControlID="txtPercentual" InputDirection="RightToLeft">
                                            </ajaxToolkit:MaskedEditExtender>
                                            <asp:TextBox ID="txtPercentual" runat="server" Width="80px"/>
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Indice">
                                        <ItemTemplate>
                                            <asp:TextBox ID="txtIndice" runat="server" ReadOnly="True" Width="56px"/>
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Desconto">
                                        <ItemTemplate>
                                            <asp:TextBox ID="txtDesconto" runat="server" ReadOnly="True" Width="80px"/>
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                </Columns>
                                <PagerStyle BackColor="White" ForeColor="#000066" HorizontalAlign="Left" />
                                <SelectedRowStyle BackColor="#669999" Font-Bold="True" ForeColor="White" />
                                <HeaderStyle BackColor="#006699" Font-Bold="True" ForeColor="White" />
                            </asp:GridView>
                            &nbsp;
                            <br />
                            <strong>Observação:</strong><br />
                            <asp:TextBox ID="txtObservacao" runat="server" Height="78px" TextMode="MultiLine"
                                Width="434px"/>
                            <br />
                            <table border="0" width="434px">
                                <tr>
                                    <td bgcolor="#FFFFC0">
                                        <asp:Button ID="btnCalcular" runat="server" CssClass="botao" OnClick="cmdCalcular_Click"
                                            Text="Calcular" 
                                            
                                            UseSubmitBehavior="False" />&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                        <asp:Button ID="btnOK" runat="server" CssClass="botao" Text="OK" OnClick="btnOK_Click"
                                            
                                            
                                            UseSubmitBehavior="False" />&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                        <asp:Button ID="btnLimpar" runat="server" CssClass="botao" Text="Limpar" OnClick="btnLimpar_Click"
                                            
                                            
                                            UseSubmitBehavior="False" />
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td align="right" class="primario">
                            Classificação:
                        </td>
                        <td>
                            <asp:DropDownList ID="DdlClassificacao" runat="server" 
                                Width="150px">
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <td align="right">
                            Pesagem:
                        </td>
                        <td style="height: 23px">
                            <asp:TextBox ID="txtPrimeiraPesagem" CssClass="txtDecimal" runat="server" Font-Names="Arial" Font-Size="9pt"
                                TabIndex="160" Width="80px"/>
                        </td>
                    </tr>
                    <tr>
                        <td align="right">
                            Peso Bruto:
                        </td>
                        <td>
                            <asp:TextBox ID="txtPesoBruto" runat="server" Enabled="False" Font-Names="Arial"
                                Font-Size="9pt" Width="80px"/>
                            <ajaxToolkit:MaskedEditExtender ID="mkePesoBruto" runat="server" TargetControlID="txtPesoBruto"
                                InputDirection="RightToLeft" Mask="999,999,999,999" MaskType="Number">
                            </ajaxToolkit:MaskedEditExtender>
                        </td>
                    </tr>
                    <tr>
                        <td align="right">
                            Desconto:
                        </td>
                        <td>
                            <asp:TextBox ID="txtDesconto" runat="server" Enabled="False" Font-Names="Arial" Font-Size="9pt"
                                Width="80px"/>
                            <ajaxToolkit:MaskedEditExtender ID="mkeDesconto" runat="server" TargetControlID="txtDesconto"
                                InputDirection="RightToLeft" Mask="999,999,999,999" MaskType="Number">
                            </ajaxToolkit:MaskedEditExtender>
                        </td>
                    </tr>
                    <tr>
                        <td align="right" class="rotulo" style="color: red">
                            Liquido:
                        </td>
                        <td>
                            <asp:TextBox ID="txtLiquido" runat="server" Enabled="False" Font-Names="Arial" Font-Size="9pt"
                                Width="80px"/>
                            <ajaxToolkit:MaskedEditExtender ID="mkeLiquido" runat="server" TargetControlID="txtLiquido"
                                InputDirection="RightToLeft" Mask="999,999,999,999" MaskType="Number">
                            </ajaxToolkit:MaskedEditExtender>
                        </td>
                    </tr>
                </table>
            </ContentTemplate>
        </asp:UpdatePanel>
    </div>
    </form>
</body>
</html>
